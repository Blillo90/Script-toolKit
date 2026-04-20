#region ═══════════════════════════════════════════════════════════
# HELPERS REMOTOS REUTILIZABLES
#═══════════════════════════════════════════════════════════════════

# ── Devuelve $true si el nombre apunta al equipo local ────────────────────────
function Test-IsLocal {
    param([string]$ComputerName)
    $n = $ComputerName.Trim().ToUpper()
    return ($n -eq $env:COMPUTERNAME.ToUpper()) -or
           ($n -eq 'LOCALHOST') -or
           ($n -eq '127.0.0.1')
}

# ── Ejecuta un scriptblock en local o remoto segun el objetivo ────────────────
# Local  : & $ScriptBlock (sin WinRM, sin red)
# Remoto : Invoke-Command con timeout via $script:RemoteSessionOpt
# ArgumentList se pasa igual en ambos casos (splatting posicional).
# OperationTimeoutMs: si >0 construye un PSSessionOption ad-hoc con ese timeout.
# Usar solo para operaciones largas (ej: inscripcion de certificados via certreq).
function Invoke-LocalOrRemote {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList       = @(),
        [int]$OperationTimeoutMs      = 0
    )
    if (Test-IsLocal $ComputerName) {
        if ($ArgumentList.Count -gt 0) { return & $ScriptBlock @ArgumentList }
        else                           { return & $ScriptBlock }
    }
    $sessOpt = if ($OperationTimeoutMs -gt 0) {
        New-PSSessionOption -OpenTimeout 10000 -OperationTimeout $OperationTimeoutMs
    } else {
        $script:RemoteSessionOpt
    }
    $opts = @{
        ComputerName  = $ComputerName
        ScriptBlock   = $ScriptBlock
        SessionOption = $sessOpt
        ErrorAction   = 'Stop'
    }
    if ($ArgumentList.Count -gt 0) { $opts['ArgumentList'] = $ArgumentList }
    return Invoke-Command @opts
}

# ── Ejecuta gpupdate /force en equipo remoto, devuelve @{Status;Details} ──────
function Invoke-RemoteGpupdate {
    param([Parameter(Mandatory)][string]$ComputerName)
    try {
        $ec = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
            (Start-Process gpupdate.exe -ArgumentList "/force /wait:0" -Wait -PassThru).ExitCode
        }
        if ($ec -eq 0) { return @{ Status = "OK";   Details = "ExitCode=0" } }
        else           { return @{ Status = "WARN"; Details = "ExitCode=$ec" } }
    } catch {
        return @{ Status = "ERROR"; Details = $_.Exception.Message }
    }
}

# ── Comprueba que CcmExec esta Running y el namespace ClientSDK accesible ─────
#    Devuelve "OK" o un string "ERROR: ..." / "VPN: ..." para mostrar en GUI
function Test-RemoteSccmReady {
    param([Parameter(Mandatory)][string]$ComputerName)
    try {
        return Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
            $svc = Get-Service "CcmExec" -ErrorAction SilentlyContinue
            if (-not $svc)                 { return "ERROR: CcmExec no instalado" }
            if ($svc.Status -ne "Running") { return "ERROR: CcmExec $($svc.Status)" }
            $ns = Get-CimInstance -Namespace "root\ccm" -ClassName "__NAMESPACE" `
                                  -Filter "Name='ClientSDK'" -ErrorAction SilentlyContinue
            if (-not $ns)                  { return "ERROR: namespace root\ccm\ClientSDK inaccesible (WMI?)" }
            return "OK"
        }
    } catch {
        # Diferenciar errores de conectividad de red (WinRM, RPC, timeout) de otros errores
        $msg = $_.Exception.Message
        if ($msg -match 'WinRM|WSMan|WS-Man|connect|network|timeout|refused|RPC|access.?denied|firewall' ) {
            return "ERROR: Sin acceso remoto (WinRM/red bloqueado). Posible VPN o firewall."
        }
        return "ERROR: $_"
    }
}

# ── Detecta la zona de red del equipo basandose en su IP resuelta ──────────────
# Devuelve: 'LOCAL' | 'VPN' | 'LAN'
# 'VPN' si la primera IPv4 resuelta esta en los rangos VPN corporativos definidos.
# Usa [System.Net.Dns]::GetHostAddresses (pila completa del OS, sin cache propio).
function Get-TargetNetworkZone {
    param([string]$Hostname)
    if (Test-IsLocal $Hostname) { return 'LOCAL' }
    try {
        $ip = [System.Net.Dns]::GetHostAddresses($Hostname) |
              Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
              Select-Object -First 1 |
              ForEach-Object { $_.ToString() }
        if ($ip -and ($ip.StartsWith('10.142.') -or $ip.StartsWith('10.99.'))) { return 'VPN' }
    } catch {}
    return 'LAN'
}

# ── Lanza los ciclos SCCM estandar (scriptblock ejecutado en equipo remoto) ───
$script:SccmCyclesBlock = {
    try {
    # Detecta si el scriptblock se ejecuta en una sesion remota (via Invoke-Command).
    # $PSSenderInfo solo esta definido en sesiones remotas; en local es $null.
    # Cuando es remota no intentamos enumerar ciclos (inestable via WinRM/CIM);
    # el disparo directo por ID sigue funcionando y es el unico camino soportado.
    $isRemoteSession = $null -ne $PSSenderInfo

    # Validacion de salud del cliente SCCM (reemplaza la enumeracion no fiable de ciclos).
    # Se reportan tres indicadores independientes:
    #   - SCCM client accessible: CcmExec instalado y Running
    #   - Client health status:   namespace root\ccm accesible via CIM
    #   - Policy trigger capability: clase SMS_Client/TriggerSchedule disponible
    $svc = Get-Service "CcmExec" -ErrorAction SilentlyContinue
    $clientAccessible = ($svc -and $svc.Status -eq 'Running')
    if (-not $svc) {
        return @{ Status = "WARN"; Details = "SCCM client accessible: NO (CcmExec no instalado)" }
    }
    if ($svc.Status -ne "Running") {
        return @{ Status = "WARN"; Details = "SCCM client accessible: NO (CcmExec $($svc.Status))" }
    }

    # Client health status: namespace root\ccm accesible.
    $namespaceOk = $false
    try {
        $null = Get-CimInstance -Namespace 'root\ccm' -ClassName '__NAMESPACE' -ErrorAction Stop
        $namespaceOk = $true
    } catch {}

    # Policy trigger capability: clase SMS_Client (TriggerSchedule) disponible.
    $triggerCapable = $false
    try {
        $null = Get-CimClass -Namespace 'root\ccm' -ClassName 'SMS_Client' -ErrorAction Stop
        $triggerCapable = $true
    } catch {}

    $health = @(
        "SCCM client accessible: YES"
        "Client health status: $(if ($namespaceOk)     { 'OK' }        else { 'ERROR (root\ccm inaccesible)' })"
        "Policy trigger capability: $(if ($triggerCapable) { 'AVAILABLE' } else { 'UNAVAILABLE' })"
    )
    if ($isRemoteSession) {
        $health += "Remote SCCM cycle enumeration is not supported - triggering cycles directly"
    }

    if (-not $triggerCapable) {
        return @{ Status = "ERROR"; Details = ($health -join " | ") }
    }

    # Ciclos objetivo identificados por ScheduleID (estable, sin dependencia de idioma).
    # SoftFail=true: el ciclo puede lanzar excepcion COM en clientes sanos
    # (sin despliegues activos, agente ocupado, sin sesion de usuario) -> se degrada a WARN, no ERROR.
    # NOTA: No se enumeran ciclos via CCM_Scheduler_ScheduledMessage (poco fiable, inconsistente
    # entre versiones de cliente y sin soporte remoto estable). Se dispara cada ciclo directamente
    # y se captura el fallo por ciclo individual.
    $targets = @(
        @{ Name="App Deployment Evaluation";   Id="{00000000-0000-0000-0000-000000000121}"; SoftFail=$true  },
        @{ Name="Discovery Data Collection";   Id="{00000000-0000-0000-0000-000000000003}"; SoftFail=$false },
        @{ Name="Hardware Inventory";          Id="{00000000-0000-0000-0000-000000000001}"; SoftFail=$false },
        @{ Name="Machine Policy Retrieval";    Id="{00000000-0000-0000-0000-000000000021}"; SoftFail=$false },
        @{ Name="Machine Policy Evaluation";   Id="{00000000-0000-0000-0000-000000000022}"; SoftFail=$false },
        @{ Name="User Policy Retrieval";       Id="{00000000-0000-0000-0000-000000000027}"; SoftFail=$true  },
        @{ Name="Software Inventory";          Id="{00000000-0000-0000-0000-000000000002}"; SoftFail=$false },
        @{ Name="SW Update Deployment Eval";   Id="{00000000-0000-0000-0000-000000000114}"; SoftFail=$true  },
        @{ Name="Software Update Scan";        Id="{00000000-0000-0000-0000-000000000113}"; SoftFail=$false },
        @{ Name="State Message Refresh";       Id="{00000000-0000-0000-0000-000000000111}"; SoftFail=$false }
    )

    $log = @() + $health
    $anyError = $false; $anyWarn = $false

    foreach ($t in $targets) {
        try {
            $rv = [int](Invoke-WmiMethod -Namespace "root\ccm" -Class "SMS_Client" `
                        -Name "TriggerSchedule" -ArgumentList @($t.Id)).ReturnValue
            if ($rv -eq 0) { $log += "$($t.Name)=OK" }
            else            { $log += "$($t.Name)=WARN(rv=$rv)"; $anyWarn = $true }
        } catch {
            $msg = $_.Exception.Message
            if ($t.SoftFail) {
                $log += "$($t.Name)=WARN($msg)"; $anyWarn = $true
            } else {
                $log += "$($t.Name)=ERROR($msg)"; $anyError = $true
            }
        }
    }

    $s = if ($anyError) { "ERROR" } elseif ($anyWarn) { "WARN" } else { "OK" }
    return @{ Status=$s; Details=($log -join " | ") }
    } catch {
        return @{ Status="ERROR"; Details="Excepcion inesperada en SccmCyclesBlock: $($_.Exception.Message)" }
    }
}

# ── Espera un job manteniendo la GUI viva via DoEvents ────────────────────────
# Devuelve $true si el job completo en plazo; $false si timeout (job ya cancelado).
# Se usa para operaciones del sistema de duracion variable (DISM, SFC, ChkDsk).
function Wait-JobWithEvents {
    param(
        [Parameter(Mandatory)]$Job,
        [int]$TimeoutMinutes   = 45,
        [string]$ProgressLabel = "Operacion",
        [int]$ProgressFrom     = -1,
        [int]$ProgressTo       = -1
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $started  = Get-Date
    $lastPing = $started

    while ($Job.State -eq 'Running' -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.Application]::DoEvents()

        # Progreso interpolado por tiempo: avanza rapido al inicio, lento al final (easing cuadratico)
        if ($ProgressFrom -ge 0 -and $ProgressTo -gt $ProgressFrom -and $script:progressBar) {
            $fraction = [Math]::Min(0.95, ((Get-Date) - $started).TotalMinutes / $TimeoutMinutes)
            $adj      = 1.0 - [Math]::Pow(1.0 - $fraction, 2)
            $val      = [int]($ProgressFrom + ($ProgressTo - $ProgressFrom) * $adj)
            if ($script:progressBar.Value -ne $val) { $script:progressBar.Value = $val }
        }

        if (((Get-Date) - $lastPing).TotalSeconds -ge 30) {
            $elapsed  = [int]((Get-Date) - $started).TotalSeconds
            Write-Info ("  ... {0} en progreso ({1}s transcurridos)..." -f $ProgressLabel, $elapsed)
            $script:outputBox.Update()
            $lastPing = Get-Date
        }
    }

    if ($Job.State -eq 'Running') {
        Stop-Job  $Job
        Remove-Job $Job -Force
        return $false
    }
    return $true
}

# ── Inscribe un certificado via certreq + CES Kerberos (3 pasos) ──────────────
#    Se ejecuta REMOTAMENTE. Devuelve @{Status;Details;Steps}
$script:CertreqEnrollBlock = {
    param([string[]]$CesUrls, [string]$CertType)
    $steps      = @()
    $cesTimeout = 20
    $rand       = Get-Random -Maximum 99999
    $base       = "$env:TEMP\AirbusEnroll_$rand"
    $infPath    = "$base.inf"
    $reqPath    = "$base.req"
    $cerPath    = "$base.cer"

    # INF minimo - sujeto lo rellena la CA desde AD
    @"
[Version]
Signature="`$Windows NT`$"

[NewRequest]
Subject = "CN=$env:COMPUTERNAME"
MachineKeySet = TRUE
KeySpec       = AT_KEYEXCHANGE
KeyLength     = 2048
Exportable    = FALSE
RequestType   = PKCS10

"@ | Out-File $infPath -Encoding ASCII

    # PASO 1: certreq -new (genera CSR localmente)
    $newOut = certreq -new -machine -q $infPath $reqPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        Remove-Item $infPath -Force -ErrorAction SilentlyContinue
        $steps += @{ Name="PASO 1: certreq -new (genera CSR)"; Status="ERROR"
                     Details="ExitCode=$($LASTEXITCODE): $($newOut -join ' ')" }
        return @{ Status="ERROR"; Details="certreq -new fallo (ExitCode=$($LASTEXITCODE))"; Steps=$steps }
    }
    $steps += @{ Name="PASO 1: certreq -new (genera CSR)"; Status="OK"; Details=$reqPath }

    # PASO 2: certreq -submit a cada CES (timeout configurable)
    $lastErr = ""; $submitOk = $false; $usedUrl = ""
    foreach ($url in $CesUrls) {
        $job = Start-Job -ScriptBlock {
            param($u, $rq, $cp)
            certreq -submit -config $u `
                -attrib "CertificateTemplate:Airbusauto-enrolledclientauthentication" `
                -q -AdminForceMachine -Kerberos $rq $cp 2>&1
        } -ArgumentList $url, $reqPath, $cerPath

        $null = Wait-Job $job -Timeout $cesTimeout
        if ($job.State -eq 'Running') {
            Stop-Job $job; Remove-Job $job -Force
            Remove-Item $cerPath -Force -ErrorAction SilentlyContinue
            $lastErr = "$url TIMEOUT (>${cesTimeout}s)"
            continue
        }
        $submitOut  = Receive-Job $job
        $submitExit = if ($job.ChildJobs[0].Error.Count -gt 0) { 1 } else { 0 }
        Remove-Job $job -Force

        if ($submitExit -eq 0 -and (Test-Path $cerPath)) {
            $submitOk = $true; $usedUrl = $url
            $steps += @{ Name="PASO 2: certreq -submit"; Status="OK"
                         Details="$($submitOut -join ' ') | via $url" }
            break
        }
        $lastErr = "$url ExitCode=$submitExit : $($submitOut -join ' ')"
        Remove-Item $cerPath -Force -ErrorAction SilentlyContinue
    }

    if (-not $submitOk) {
        Remove-Item $infPath, $reqPath -Force -ErrorAction SilentlyContinue
        $steps += @{ Name="PASO 2: certreq -submit"; Status="ERROR"
                     Details="Todos los CES fallaron. Ultimo: $lastErr" }
        return @{ Status="ERROR"; Details="certreq -submit fallo en todos los CES"; Steps=$steps }
    }

    # PASO 3: certreq -accept (instala el certificado)
    $acceptOut = certreq -accept -machine $cerPath 2>&1
    Remove-Item $infPath, $reqPath, $cerPath -Force -ErrorAction SilentlyContinue

    if ($LASTEXITCODE -eq 0) {
        $steps += @{ Name="PASO 3: certreq -accept (instala cert)"; Status="OK"
                     Details="Instalado en LocalMachine\My" }
        return @{ Status="OK"; Details="$CertType inscrito OK via $usedUrl"; Steps=$steps }
    }
    $steps += @{ Name="PASO 3: certreq -accept (instala cert)"; Status="WARN"
                 Details="$($acceptOut -join ' ')" }
    return @{ Status="WARN"; Details="Emitido pero accept fallo: $($acceptOut -join ' ')"; Steps=$steps }
}

#endregion
