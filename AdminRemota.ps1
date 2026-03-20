#Requires -Version 5.1
<#
.SYNOPSIS
    Herramienta de administracion remota unificada v2.8.3 (GUI)
.DESCRIPTION
    Interfaz grafica con opciones de administracion remota:
      1. Comprobar Masterizacion de un equipo
      2. Comprobar Software en Centro de Software (SCCM)
      3. Borrar Drivers USB de un equipo
      4. Informacion del Sistema
.AUTHOR
    Pablo Perez Herrero - AC1974
.COMPANYNAME
    Accenture
.VERSION
    2.8.6
#>

[CmdletBinding()]
param(
    # Patron de issuer para validar el certificado en modo Nacional.
    # Solo se usa en Invoke-MasterCheck rama Nacional; modo Divisional usa $script:DivisionalCerts.
    # Como parametro de script-level, es accesible en funciones como $script:ExpectedIssuerLike.
    [string]$ExpectedIssuerLike = "*Airbus Issuing CA Juan de la Cierva*"
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Ocultar la ventana de consola: herramienta GUI pura.
Add-Type -Name ConsoleHelper -Namespace Win32 -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
[Win32.ConsoleHelper]::ShowWindow([Win32.ConsoleHelper]::GetConsoleWindow(), 0) | Out-Null

$ErrorActionPreference = "Continue"
$PSDefaultParameterValues['Invoke-Command:ErrorAction']  = 'SilentlyContinue'
$PSDefaultParameterValues['Get-CimInstance:ErrorAction'] = 'SilentlyContinue'

#region ═══════════════════════════════════════════════════════════
# CONSTANTES Y ESTADO GLOBAL
#═══════════════════════════════════════════════════════════════════

$script:outputBox       = $null
$script:statusLabel     = $null
$script:progressBar     = $null
$script:progressLabel   = $null
$script:cancelRequested = $false
$script:EquipoSelCard   = $null        # tarjeta actualmente seleccionada en el panel lateral
$script:Modo            = "Nacional"   # "Nacional" | "Divisional"
$script:Target          = ""
$script:StepResults     = New-Object System.Collections.Generic.List[object]

# Colores GUI (definidos aqui para que esten disponibles en toda la sesion)
$script:White  = [System.Drawing.Color]::White
$script:Silver = [System.Drawing.Color]::Silver

# Configuracion de certificados Divisional: unica fuente de verdad.
# Cada entrada define el nombre del cert, el filtro de issuer para deteccion y las URLs CES.
# El nombre es la clave de union entre deteccion y enrollment — no duplicar en otro lugar.
$script:DivisionalCerts = @(
    @{
        Name    = "Breguet G1"
        Filter  = "*Breguet*"
        CesUrls = @(
            "https://aefews01.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20Breguet%20G1_CES_Kerberos/service.svc/CES",
            "https://aefews02.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20Breguet%20G1_CES_Kerberos/service.svc/CES"
        )
    },
    @{
        Name    = "da Vinci G1"
        Filter  = "*Vinci*"
        CesUrls = @(
            "https://aefews01.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20da%20Vinci%20G1_CES_Kerberos/service.svc/CES",
            "https://aefews02.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20da%20Vinci%20G1_CES_Kerberos/service.svc/CES"
        )
    }
)

#endregion

#region ═══════════════════════════════════════════════════════════
# FUNCIONES DE SALIDA GUI
#═══════════════════════════════════════════════════════════════════

function Append-Output {
    param([string]$Text, [System.Drawing.Color]$Color)
    if (-not $script:outputBox) { return }
    $script:outputBox.SelectionStart  = $script:outputBox.TextLength
    $script:outputBox.SelectionLength = 0
    $script:outputBox.SelectionColor  = $Color
    $script:outputBox.AppendText("$Text`r`n")
    $script:outputBox.SelectionStart  = $script:outputBox.TextLength
    $script:outputBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Write-Info { param($Msg) Append-Output $Msg                 ([System.Drawing.Color]::Cyan)                   }
function Write-Ok   { param($Msg) Append-Output "  OK: $Msg"        ([System.Drawing.Color]::LightGreen)             }
function Write-Warn { param($Msg) Append-Output "  WARN: $Msg"      ([System.Drawing.Color]::Yellow)                 }
function Write-Fail { param($Msg) Append-Output "  ERROR: $Msg"     ([System.Drawing.Color]::Tomato)                 }
function Write-Sep  {             Append-Output ("-" * 65)           ([System.Drawing.Color]::FromArgb(80, 80, 80))   }

function Set-Status {
    param([string]$Msg, [System.Drawing.Color]$Color = [System.Drawing.Color]::White)
    if (-not $script:statusLabel) { return }
    $script:statusLabel.Text      = "  $Msg"
    $script:statusLabel.ForeColor = $Color
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Progress {
    param([int]$Value, [string]$Label = "")
    if ($script:progressBar) {
        $script:progressBar.Value = [Math]::Max(0, [Math]::Min(100, $Value))
    }
    if ($script:progressLabel -and $Label -ne "") {
        $script:progressLabel.Text      = $Label
        $script:progressLabel.ForeColor = if ($Value -ge 100) {
            [System.Drawing.Color]::LightGreen
        } else {
            $script:Silver
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Confirm-Action {
    param([string]$Message, [string]$Title = "Confirmar")
    $r = [System.Windows.Forms.MessageBox]::Show(
        $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Get-Input {
    param([string]$Prompt, [string]$Title = "Entrada requerida", [string]$Default = "")
    return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $Default)
}

# Escribe en GUI los sub-pasos devueltos por operaciones remotas (array de @{Name;Status;Details})
function Write-StepList {
    param([array]$Steps)
    foreach ($s in $Steps) {
        if (-not $s) { continue }
        switch ($s.Status) {
            "OK"    { Write-Ok   "  $($s.Name)  ->  $($s.Details)" }
            "WARN"  { Write-Warn "  $($s.Name)  ->  $($s.Details)" }
            "ERROR" { Write-Fail "  $($s.Name)  ->  $($s.Details)" }
        }
    }
}

#endregion

#region ═══════════════════════════════════════════════════════════
# LOGICA COMUN: StepResults y orquestacion
#═══════════════════════════════════════════════════════════════════

function Reset-StepResults {
    $script:StepResults = New-Object System.Collections.Generic.List[object]
}

function Add-StepResult {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][ValidateSet("OK","WARN","ERROR")][string]$Status,
        [string]$Details = ""
    )
    $script:StepResults.Add([PSCustomObject]@{
        Step    = $Step
        Status  = $Status
        Details = $Details
        Time    = (Get-Date)
    }) | Out-Null
}

function Invoke-Step {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )
    if ($script:cancelRequested) {
        Write-Warn "Cancelado antes de: $Name"
        return
    }
    Append-Output "" $script:White
    Write-Info "---- $Name ----"
    try {
        $res     = & $ScriptBlock
        # Si el scriptblock devuelve null (Invoke-Command silenciado por red/acceso), se
        # registra como WARN en lugar de OK para no dar falsos positivos en el resumen.
        $status  = if     ($res -is [hashtable] -and $res.Status) { $res.Status }
                   elseif ($null -eq $res)                         { "WARN"      }
                   else                                            { "OK"        }
        $details = if ($res -is [hashtable] -and $res.Details) { ($res.Details | Out-String).Trim() }
                   elseif ($null -eq $res)                      { "Sin respuesta remota (red o acceso?)" }
                   else                                         { "" }
        Add-StepResult -Step $Name -Status $status -Details $details
        switch ($status) {
            "OK"    { Write-Ok   $Name }
            "WARN"  { Write-Warn "$Name | $details" }
            "ERROR" { Write-Fail "$Name | $details" }
        }
    } catch {
        Add-StepResult -Step $Name -Status "ERROR" -Details $_.Exception.Message
        Write-Fail "$Name | $($_.Exception.Message)"
    }
}

function Show-Summary {
    param([string]$ExcludeWarnStep = "")
    Append-Output "" $script:White
    Write-Sep
    Write-Info "RESUMEN FINAL"
    Write-Sep
    foreach ($r in $script:StepResults) {
        $line = "  {0,-40} [{1}]  {2}" -f $r.Step, $r.Status, $r.Details
        switch ($r.Status) {
            "OK"    { Append-Output $line ([System.Drawing.Color]::LightGreen) }
            "WARN"  { Append-Output $line ([System.Drawing.Color]::Yellow)     }
            "ERROR" { Append-Output $line ([System.Drawing.Color]::Tomato)     }
        }
    }
    $blocked = $script:StepResults | Where-Object {
        $_.Status -eq "ERROR" -or
        ($_.Status -eq "WARN" -and ($ExcludeWarnStep -eq "" -or $_.Step -ne $ExcludeWarnStep))
    }
    Append-Output "" $script:White
    if (-not $blocked) { Write-Ok   "TODO OK" }
    else                { Write-Warn "Hay pasos con WARN/ERROR que requieren atencion." }
}

#endregion

#region ═══════════════════════════════════════════════════════════
# HELPERS REMOTOS REUTILIZABLES
#═══════════════════════════════════════════════════════════════════

# ── Ejecuta gpupdate /force en equipo remoto, devuelve @{Status;Details} ──────
function Invoke-RemoteGpupdate {
    param([Parameter(Mandatory)][string]$ComputerName)
    $ec = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        (Start-Process gpupdate.exe -ArgumentList "/force /wait:0" -Wait -PassThru).ExitCode
    }
    if ($ec -eq 0) { return @{ Status = "OK";   Details = "ExitCode=0" } }
    else            { return @{ Status = "WARN"; Details = "ExitCode=$ec" } }
}

# ── Comprueba que CcmExec esta Running y el namespace ClientSDK accesible ─────
#    Devuelve "OK" o un string "ERROR: ..." para mostrar en GUI
function Test-RemoteSccmReady {
    param([Parameter(Mandatory)][string]$ComputerName)
    return Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        $svc = Get-Service "CcmExec" -ErrorAction SilentlyContinue
        if (-not $svc)                 { return "ERROR: CcmExec no instalado" }
        if ($svc.Status -ne "Running") { return "ERROR: CcmExec $($svc.Status)" }
        $ns = Get-CimInstance -Namespace "root\ccm" -ClassName "__NAMESPACE" `
                              -Filter "Name='ClientSDK'" -ErrorAction SilentlyContinue
        if (-not $ns)                  { return "ERROR: namespace root\ccm\ClientSDK inaccesible (WMI?)" }
        return "OK"
    }
}

# ── Lanza los ciclos SCCM estandar (scriptblock ejecutado en equipo remoto) ───
$script:SccmCyclesBlock = {
    $actions = @(
        @{ Name="App Deployment Evaluation";   Id="{00000000-0000-0000-0000-000000000121}" },
        @{ Name="Discovery Data Collection";   Id="{00000000-0000-0000-0000-000000000003}" },
        @{ Name="Hardware Inventory";          Id="{00000000-0000-0000-0000-000000000001}" },
        @{ Name="Machine Policy Retrieval";    Id="{00000000-0000-0000-0000-000000000021}" },
        @{ Name="Machine Policy Evaluation";   Id="{00000000-0000-0000-0000-000000000022}" },
        @{ Name="Software Inventory";          Id="{00000000-0000-0000-0000-000000000002}" },
        @{ Name="SW Update Deployment Eval";   Id="{00000000-0000-0000-0000-000000000114}" },
        @{ Name="Software Update Scan";        Id="{00000000-0000-0000-0000-000000000113}" },
        @{ Name="State Message Refresh";       Id="{00000000-0000-0000-0000-000000000111}" }
    )
    $log = @(); $anyError = $false; $anyWarn = $false
    foreach ($a in $actions) {
        try {
            $rv = [int](Invoke-WmiMethod -Namespace "root\ccm" -Class "SMS_Client" `
                        -Name "TriggerSchedule" -ArgumentList @($a.Id)).ReturnValue
            if ($rv -eq 0) { $log += "$($a.Name)=OK" }
            else            { $log += "$($a.Name)=WARN($rv)"; $anyWarn = $true }
        } catch {
            $log += "$($a.Name)=ERROR($($_.Exception.Message))"; $anyError = $true
        }
    }
    $s = if ($anyError) { "ERROR" } elseif ($anyWarn) { "WARN" } else { "OK" }
    return @{ Status=$s; Details=($log -join " | ") }
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

#region ═══════════════════════════════════════════════════════════
# OPCION 1 - COMPROBAR MASTERIZACION
#═══════════════════════════════════════════════════════════════════

function Invoke-MasterCheck {
    param([Parameter(Mandatory)][string]$ComputerName)
    # $script:Target captura el equipo en scope de script para que los scriptblocks
    # pasados a Invoke-Step puedan acceder a el sin depender de la captura lexical
    # de closures anidados, que es fragil en PS5.1 con scriptblocks pasados como parametro.
    $script:Target = $ComputerName
    Reset-StepResults

    Write-Info "Masterizacion de '$ComputerName'  [Modo: $($script:Modo)]"
    Write-Sep

    # Step 1: GPUpdate
    Invoke-Step -Name "GPUPDATE /force" -ScriptBlock {
        Invoke-RemoteGpupdate -ComputerName $script:Target
    }

    # Step 2: Certificados (logica diferente segun modo)
    if ($script:Modo -eq "Divisional") {
        Invoke-Step -Name "Certificados Divisional (Breguet G1 + da Vinci G1)" -ScriptBlock {

            # Pasar solo Name y Filter al remoto (CesUrls no son necesarias para deteccion)
            $certDefsForRemote = @($script:DivisionalCerts | ForEach-Object { @{ Name=$_.Name; Filter=$_.Filter } })
            $result = Invoke-Command -ComputerName $script:Target -ArgumentList (,$certDefsForRemote) -ScriptBlock {
                param([object[]]$certDefs)
                $all     = Get-ChildItem "Cert:\LocalMachine\My" -ErrorAction SilentlyContinue
                $now     = Get-Date
                $details = @(); $missing = @(); $cleaned = @()

                foreach ($caEntry in $certDefs) {
                    $certs = @($all | Where-Object {
                        $_.Issuer -like $caEntry.Filter -and $_.NotAfter -gt $now
                    } | Sort-Object NotAfter -Descending)

                    if ($certs.Count -eq 0) {
                        $missing += $caEntry.Name
                    } else {
                        if ($certs.Count -gt 1) {
                            $certs | Select-Object -Skip 1 | ForEach-Object {
                                Remove-Item "Cert:\LocalMachine\My\$($_.Thumbprint)" -Force -ErrorAction SilentlyContinue
                            }
                            $cleaned += "$($caEntry.Name): $($certs.Count - 1) duplicado(s) borrado(s)"
                        }
                        $details += "$($caEntry.Name): Expira=$($certs[0].NotAfter.ToString('dd/MM/yyyy'))"
                    }
                }

                $allDetails = $cleaned + $details
                if ($missing.Count -gt 0) {
                    return @{ Status="ERROR"; Details="FALTA: $($missing -join ', ')  |  $($allDetails -join ' | ')"; Missing=$missing; Cleaned=$cleaned }
                }
                return @{ Status="OK"; Details=($allDetails -join " | "); Cleaned=$cleaned }
            }

            # Informar duplicados borrados
            foreach ($c in @($result.Cleaned)) { if ($c) { Write-Warn "  $c" } }

            # Ofrecer inscripcion via CES si faltan certs
            $missingCerts = @($result.Missing)
            if ($result.Status -eq "ERROR" -and $missingCerts.Count -gt 0) {
                $missingStr = $missingCerts -join ", "
                if (Confirm-Action "Faltan certs en '$($script:Target)': $missingStr.`nInscribir via CES Kerberos (aefews01/02)?") {
                    foreach ($certType in $missingCerts) {
                        $certDef = $script:DivisionalCerts | Where-Object { $_.Name -eq $certType } | Select-Object -First 1
                        $urls    = @($certDef.CesUrls)
                        $ct      = $certType
                        Invoke-Step -Name "Inscribir $ct via certreq+CES" -ScriptBlock {
                            $remoteResult = Invoke-Command -ComputerName $script:Target `
                                -ArgumentList $urls, $ct, $script:CertreqEnrollBlock -ScriptBlock {
                                    param($cesUrls, $certType, $enrollBlock)
                                    & $enrollBlock $cesUrls $certType
                                }
                            Write-StepList -Steps @($remoteResult.Steps)
                            return $remoteResult
                        }
                    }

                    # Si todos los certs faltantes se inscribieron OK -> WARN (corregido)
                    $okCount = @($script:StepResults | Where-Object {
                        $_.Step -like "Inscribir * via certreq+CES" -and $_.Status -eq "OK"
                    }).Count
                    if ($okCount -eq $missingCerts.Count) {
                        return @{ Status="WARN"; Details="Corregido: $missingStr inscrito(s) OK | $($result.Details)" }
                    }
                }
            }
            return $result
        }

    } else {
        # Modo Nacional: un cert con el issuer del parametro
        Invoke-Step -Name "Certificado LocalMachine\My" -ScriptBlock {
            Invoke-Command -ComputerName $script:Target -ArgumentList $script:ExpectedIssuerLike -ScriptBlock {
                param($issuer)
                $certs = Get-ChildItem "Cert:\LocalMachine\My" -ErrorAction SilentlyContinue
                if (-not $certs -or $certs.Count -eq 0) {
                    return @{ Status="ERROR"; Details="Sin certificados en LocalMachine\My" }
                }
                $status  = "OK"
                $details = @()
                if ($certs.Count -gt 1) { $details += "Hay $($certs.Count) certs (debe ser 1)"; $status = "WARN" }
                $cert = $certs[0]
                if ($cert.Issuer -notlike $issuer) { return @{ Status="ERROR"; Details="Issuer no coincide: $($cert.Issuer)" } }
                if ($cert.NotAfter -le (Get-Date))  { return @{ Status="ERROR"; Details="CADUCADO: NotAfter=$($cert.NotAfter)" } }
                $details += "Issuer=$($cert.Issuer) | NotAfter=$($cert.NotAfter)"
                return @{ Status=$status; Details=($details -join " | ") }
            }
        }
    }

    # Step 3: success.txt
    Invoke-Step -Name "success.txt" -ScriptBlock {
        Invoke-Command -ComputerName $script:Target -ScriptBlock {
            $f = Get-Item "C:\success.txt" -ErrorAction SilentlyContinue
            if (-not $f) { return @{ Status="ERROR"; Details="No existe C:\success.txt" } }
            $s = if ($f.LastWriteTime.Date -eq (Get-Date).Date) { "OK" } else { "WARN" }
            return @{ Status=$s; Details="Fecha: $($f.LastWriteTime)" }
        }
    }

    # Step 4: Ciclos SCCM
    Invoke-Step -Name "SCCM Client Cycles" -ScriptBlock {
        Invoke-Command -ComputerName $script:Target -ScriptBlock $script:SccmCyclesBlock
    }

    # Step 5: Centro de Software
    Invoke-Step -Name "Centro de Software (CCM_Application)" -ScriptBlock {
        $diag = Test-RemoteSccmReady -ComputerName $script:Target
        if ($diag -ne "OK") { return @{ Status="ERROR"; Details=$diag } }

        Invoke-Command -ComputerName $script:Target -ScriptBlock {
            $app = Get-CimInstance -Namespace "root\ccm\ClientSDK" -ClassName "CCM_Application" `
                                   -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($app) {
                $ver = (Get-CimInstance -Namespace "root\ccm" -ClassName "SMS_Client" `
                                        -ErrorAction SilentlyContinue).ClientVersion
                return @{ Status="OK"; Details="Namespace accesible | ClientVersion=$ver" }
            }
            return @{ Status="WARN"; Details="CcmExec Running y namespace OK, pero sin apps en CCM_Application" }
        }
    }

    # Step 6 y 7: Software de seguridad (patron unificado; Filters es array de globs)
    foreach ($swCheck in @(
        @{ Name="Cisco Secure Client"; Filters=@("*Cisco Secure Client*","*Cisco AnyConnect*") },
        @{ Name="Trellix / McAfee";    Filters=@("*Trellix*","*McAfee*","*Endpoint Security*") }
    )) {
        $checkName    = $swCheck.Name
        $checkFilters = $swCheck.Filters
        Invoke-Step -Name $checkName -ScriptBlock {
            $found = Invoke-Command -ComputerName $script:Target -ArgumentList (,$checkFilters) -ScriptBlock {
                param([string[]]$filters)
                $paths = @(
                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
                )
                foreach ($p in $paths) {
                    $entries = Get-ItemProperty $p -ErrorAction SilentlyContinue
                    foreach ($f in $filters) {
                        $m = $entries | Where-Object { $_.DisplayName -like $f } | Select-Object -First 1
                        if ($m) { return "$($m.DisplayName) v$($m.DisplayVersion)" }
                    }
                }
                return $null
            }
            if ($found) { return @{ Status="OK";    Details=$found } }
            else         { return @{ Status="ERROR"; Details="No encontrado en Uninstall" } }
        }
    }

    Show-Summary -ExcludeWarnStep "success.txt"
}

#endregion

#region ═══════════════════════════════════════════════════════════
# OPCION 2 - COMPROBAR SOFTWARE EN CENTRO DE SOFTWARE
#═══════════════════════════════════════════════════════════════════

function Invoke-SoftwareCheck {
    param([Parameter(Mandatory)][string]$ComputerName)
    $script:Target = $ComputerName

    $appName = Get-Input "Nombre o ID de la aplicacion a buscar" "Buscar en Centro de Software"
    if ([string]::IsNullOrWhiteSpace($appName)) { return }

    Write-Info "Buscando '$appName' en '$ComputerName'..."
    Write-Sep

    # Diagnostico previo (reutiliza helper comun)
    $diag = Test-RemoteSccmReady -ComputerName $ComputerName
    if ($diag -ne "OK") {
        Write-Fail "No se puede consultar el Centro de Software: $diag"
        return
    }

    $apps = Get-CimInstance -ComputerName $ComputerName -Namespace "root\ccm\ClientSDK" `
                            -ClassName "CCM_Application" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*$appName*" }

    if ($apps) {
        Write-Ok "Aplicacion(es) encontrada(s) en Centro de Software:"
        foreach ($a in $apps) {
            $aColor = switch ($a.InstallState) {
                "Installed"    { [System.Drawing.Color]::LightGreen  }
                "NotInstalled" { [System.Drawing.Color]::Yellow      }
                default        { [System.Drawing.Color]::LightYellow }
            }
            Append-Output ("    > {0}  |  Estado: {1}  |  Version: {2}" -f $a.Name, $a.InstallState, $a.SoftwareVersion) $aColor
        }

        $notInstalled = @($apps | Where-Object { $_.InstallState -eq "NotInstalled" })
        if ($notInstalled.Count -gt 0) {
            $appList = ($notInstalled | ForEach-Object { $_.Name }) -join "`n  - "
            if (Confirm-Action "Las siguientes aplicaciones NO estan instaladas en '$ComputerName':`n`n  - $appList`n`n¿Lanzar instalacion via SCCM ahora?") {
                foreach ($app in $notInstalled) {
                    Write-Info "Lanzando instalacion de '$($app.Name)'..."
                    try {
                        $rv = Invoke-Command -ComputerName $ComputerName `
                            -ArgumentList $app.Id, $app.Revision, $app.IsMachineTarget -ScriptBlock {
                                param($appId, $appRev, $isMachine)
                                $r = Invoke-CimMethod -Namespace "root\ccm\clientsdk" -ClassName "CCM_Application" `
                                    -MethodName "Install" -ErrorAction Stop -Arguments @{
                                        Id                = $appId
                                        Revision          = $appRev
                                        IsMachineTarget   = $isMachine
                                        EnforcePreference = [uint32]0
                                        Priority          = "High"
                                        IsRebootIfNeeded  = $false
                                    }
                                return [int]$r.ReturnValue
                            }
                        if ($rv -eq 0) {
                            Write-Ok "'$($app.Name)': instalacion iniciada."
                            Set-Status "Instalacion iniciada: $($app.Name)" ([System.Drawing.Color]::LightGreen)
                        } else {
                            Write-Warn "'$($app.Name)': Install devolvio codigo $rv."
                            Set-Status "WARN Install rv=$rv : $($app.Name)" ([System.Drawing.Color]::Yellow)
                        }
                    } catch {
                        Write-Fail "Error al instalar '$($app.Name)': $_"
                        Set-Status "Error al instalar: $($app.Name)" ([System.Drawing.Color]::Tomato)
                    }
                }
            }
        }
        return
    }

    # App no encontrada en SCCM
    Write-Warn "'$appName' no esta en el Centro de Software (CCM_Application)."
    Write-Info "  Posibles causas: app no desplegada, politica no aplicada aun, o solo desplegada a usuario."

    # Busqueda en registro local como fallback
    Write-Info "  Comprobando instalacion local en registro (Uninstall)..."
    $localMatches = Invoke-Command -ComputerName $ComputerName -ArgumentList $appName -ScriptBlock {
        param($name)
        $paths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        $found = @()
        foreach ($p in $paths) {
            $found += Get-ItemProperty $p -ErrorAction SilentlyContinue |
                      Where-Object { $_.DisplayName -like "*$name*" } |
                      Select-Object DisplayName, DisplayVersion, InstallDate, Publisher
        }
        return $found
    }

    if ($localMatches -and @($localMatches).Count -gt 0) {
        Write-Ok "  Encontrado en registro local (instalado, no visible en SCCM):"
        foreach ($m in @($localMatches)) {
            Append-Output ("    > {0}  v{1}  |  Instalado: {2}  |  {3}" -f `
                $m.DisplayName, $m.DisplayVersion, $m.InstallDate, $m.Publisher) `
                ([System.Drawing.Color]::LightYellow)
        }
    } else {
        Write-Warn "  Tampoco encontrado en registro local - no instalado en '$ComputerName'."
    }

    if (-not (Confirm-Action "Forzar ciclos SCCM en '$ComputerName'?")) { return }

    Reset-StepResults

    Invoke-Step -Name "SCCM Client Cycles" -ScriptBlock {
        Invoke-Command -ComputerName $script:Target -ScriptBlock $script:SccmCyclesBlock
    }

    if (Confirm-Action "Ejecutar gpupdate /force en '$ComputerName'?") {
        Invoke-Step -Name "GPUPDATE /force" -ScriptBlock {
            Invoke-RemoteGpupdate -ComputerName $script:Target
        }
    }

    Show-Summary
}

#endregion

#region ═══════════════════════════════════════════════════════════
# OPCION 4 - INFORMACION DEL SISTEMA
#═══════════════════════════════════════════════════════════════════

function Invoke-SystemInfo {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Info "Recopilando informacion del sistema '$ComputerName'..."
    Write-Sep

    $info = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        $os   = Get-CimInstance Win32_OperatingSystem
        $cs   = Get-CimInstance Win32_ComputerSystem
        $bios = Get-CimInstance Win32_BIOS
        $cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
        $nets = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled }
        $disks = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } |
                     Select-Object DeviceID,
                                   @{N="SizeGB"; E={[math]::Round($_.Size/1GB, 1)}},
                                   @{N="FreeGB"; E={[math]::Round($_.FreeSpace/1GB, 1)}}
        try   { $fqdn = [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName }
        catch { $fqdn = $env:COMPUTERNAME }

        $sccmMP = $null; $sccmSite = $null; $sccmVer = $null
        try {
            $sccmAuth = Get-WmiObject -Namespace root\ccm -Class SMS_Authority -ErrorAction Stop |
                            Select-Object -First 1
            $sccmMP   = $sccmAuth.CurrentManagementPoint
            $sccmReg  = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client" -ErrorAction SilentlyContinue
            $sccmSite = $sccmReg.AssignedSiteCode
            $sccmVer  = $sccmReg.ProductVersion
        } catch { }

        return @{
            OSCaption   = $os.Caption
            OSVersion   = $os.Version
            OSBuild     = $os.BuildNumber
            LastBoot    = $os.LastBootUpTime
            FQDN        = $fqdn
            Domain      = $cs.Domain
            User        = $cs.UserName
            CPU         = $cpu.Name
            CPUCores    = $cpu.NumberOfLogicalProcessors
            TotalRAMGB  = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
            FreeRAMGB   = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
            SerialBIOS  = $bios.SerialNumber
            BIOSVersion = $bios.SMBIOSBIOSVersion
            Networks    = @($nets | Select-Object Description, IPAddress, IPSubnet, DefaultIPGateway, MACAddress)
            Disks       = @($disks)
            SccmMP      = $sccmMP
            SccmSite    = $sccmSite
            SccmVer     = $sccmVer
        }
    }

    $cyan   = [System.Drawing.Color]::FromArgb(0, 190, 255)
    $yellow = [System.Drawing.Color]::LightYellow
    $white  = $script:White

    Append-Output "  [SISTEMA OPERATIVO]" $cyan
    Append-Output ("    SO:          {0}" -f $info.OSCaption) $white
    Append-Output ("    Version:     {0}  (Build {1})" -f $info.OSVersion, $info.OSBuild) $white
    Append-Output ("    Ultimo boot: {0}" -f $info.LastBoot.ToString("yyyy-MM-dd HH:mm")) $white
    Append-Output ("    FQDN:        {0}" -f $info.FQDN) $white
    Append-Output ("    Dominio:     {0}" -f $info.Domain) $white
    Append-Output ("    Usuario:     {0}" -f $(if ($info.User) { $info.User } else { "(nadie logueado)" })) $white

    Write-Sep
    Append-Output "  [HARDWARE]" $cyan
    Append-Output ("    CPU:         {0}" -f $info.CPU.Trim()) $white
    Append-Output ("    Nucleos log: {0}" -f $info.CPUCores) $white
    Append-Output ("    RAM total:   {0} GB   (libre: {1} GB)" -f $info.TotalRAMGB, $info.FreeRAMGB) $white
    Append-Output ("    S/N BIOS:    {0}" -f $info.SerialBIOS) $white
    Append-Output ("    BIOS ver:    {0}" -f $info.BIOSVersion) $white

    Write-Sep
    Append-Output "  [RED]" $cyan
    foreach ($n in $info.Networks) {
        $ips  = (@($n.IPAddress)         | Where-Object { $_ -notlike "*:*" }) -join ", "
        $mask = (@($n.IPSubnet)          | Where-Object { $_ -notlike "*:*" }) -join ", "
        $gw   = (@($n.DefaultIPGateway)) -join ", "
        Append-Output ("    [{0}]" -f $n.Description) $yellow
        Append-Output ("      IP / Mask: {0} / {1}" -f $ips, $mask) $white
        Append-Output ("      Gateway:   {0}" -f $gw) $white
        Append-Output ("      MAC:       {0}" -f $n.MACAddress) $white
    }

    Write-Sep
    Append-Output "  [DISCOS]" $cyan
    foreach ($d in $info.Disks) {
        $used  = [math]::Round($d.SizeGB - $d.FreeGB, 1)
        $pct   = if ($d.SizeGB -gt 0) { [math]::Round($used / $d.SizeGB * 100) } else { 0 }
        $dcolor = if     ($pct -gt 85) { [System.Drawing.Color]::Tomato     }
                  elseif ($pct -gt 70) { [System.Drawing.Color]::Yellow     }
                  else                 { [System.Drawing.Color]::LightGreen }
        Append-Output ("    {0}  {1} / {2} GB usados  ({3}%)" -f $d.DeviceID, $used, $d.SizeGB, $pct) $dcolor
    }

    Write-Sep
    Append-Output "  [SCCM CLIENT]" $cyan
    if ($info.SccmMP) {
        Append-Output ("    Mgmt Point:  {0}" -f $info.SccmMP)   ([System.Drawing.Color]::LightGreen)
        Append-Output ("    Site Code:   {0}" -f $(if ($info.SccmSite) { $info.SccmSite } else { "(desconocido)" })) $white
        Append-Output ("    Version:     {0}" -f $(if ($info.SccmVer)  { $info.SccmVer  } else { "(desconocida)"  })) $white
    } else {
        Append-Output "    Cliente SCCM no encontrado o no responde (root\ccm inaccesible)" ([System.Drawing.Color]::Gray)
    }

    Write-Sep
    Append-Output "" $white
}

#endregion

#region ═══════════════════════════════════════════════════════════
# OPCION 3 - BORRAR DRIVERS USB
#═══════════════════════════════════════════════════════════════════

function Invoke-UsbDriverClean {
    param([Parameter(Mandatory)][string]$ComputerName)

    # ── Fase A: Eleccion de modo ─────────────────────────────────────
    Write-Sep
    Write-Info "Borrado de drivers USB en '$ComputerName'"
    Write-Sep
    Append-Output "" $script:White

    # Paso 1: preguntar por drivers fantasma
    $step1 = [System.Windows.Forms.MessageBox]::Show(
        "Desea borrar drivers USB fantasma (desconectados / no presentes)?",
        "Borrado USB - Paso 1 de 2",
        [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($step1 -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Warn "Operacion cancelada por el usuario."
        return
    }

    $soloFantasmas = $false
    if ($step1 -eq [System.Windows.Forms.DialogResult]::Yes) {
        $soloFantasmas = $true
    } else {
        # Paso 2: preguntar por todos los drivers USB
        $step2 = [System.Windows.Forms.MessageBox]::Show(
            "Desea borrar TODOS los drivers USB (presentes y no presentes)?",
            "Borrado USB - Paso 2 de 2",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($step2 -ne [System.Windows.Forms.DialogResult]::Yes) {
            Write-Warn "Operacion cancelada por el usuario."
            return
        }
    }

    $modoTexto = if ($soloFantasmas) { "Solo fantasmas (desconectados)" } else { "Todos los drivers USB" }

    Write-Info "Modo seleccionado: $modoTexto"
    Append-Output "" $script:White

    # ── Fase B: Deteccion ────────────────────────────────────────────
    Write-Info "Buscando drivers candidatos en '$ComputerName'..."

    $remoteData = Invoke-Command -ComputerName $ComputerName -ArgumentList $soloFantasmas -ScriptBlock {
        param([bool]$onlyGhost)
        # Get-PnpDevice SIN parametros devuelve todos los dispositivos (presentes + fantasma).
        # -PresentOnly es el switch que RESTRINGE a solo presentes; sin el, se obtiene todo.
        # NUNCA usar "-PresentOnly $false": SwitchParameter sin ":" activa el switch y pasa
        # $false como argumento posicional, haciendo que Get-PnpDevice devuelva array vacio.
        $all      = Get-PnpDevice
        $excluded = @()

        if ($onlyGhost) {
            # Fantasmas: Status=Unknown identifica dispositivos ausentes/desconectados.
            # Se busca por FriendlyName O Class para no perder dispositivos USB cuya clase
            # sea HIDClass, DiskDrive, etc. (no todos los USB tienen Class="USB").
            $all = $all | Where-Object {
                ($_.FriendlyName -match "USB" -or $_.Class -match "USB") -and $_.Status -eq "Unknown"
            }
        } else {
            # ── Exclusiones obligatorias en modo "todos los USB" ──────────────────
            # Razon: pnputil /remove-device sobre ciertos dispositivos bloquea la
            # sesion remota o congela la operacion indefinidamente.

            # 1) USB Root Hub: controladores de bus de nivel superior.
            #    Intentar eliminarlos desconecta todos los hijos del bus (incluido
            #    el adaptador de red que mantiene WinRM) → bloqueo indefinido.
            # 2) Adaptadores de red USB (Realtek USB, ASIX USB, USB Ethernet, etc.):
            #    Eliminarlos corta la conexion de red en mitad del Invoke-Command
            #    → la sesion WinRM colapsa con el job a medias → bloqueo/error.
            #    El script de referencia los separa explicitamente como "NotSafe".
            $excluded = @($all | Where-Object {
                $_.Class -match "USB" -and (
                    $_.FriendlyName -match "Root Hub" -or
                    $_.FriendlyName -match "Realtek USB" -or
                    $_.FriendlyName -match "ASIX USB" -or
                    $_.FriendlyName -match "USB Ethernet" -or
                    $_.FriendlyName -match "USB.*Network" -or
                    $_.FriendlyName -match "USB.*LAN"
                )
            } | Sort-Object FriendlyName |
              Select-Object Status, Class, FriendlyName, InstanceId,
                @{ N="ExcludeReason"; E={
                    if ($_.FriendlyName -match "Root Hub") { "Controlador de bus (Root Hub)" }
                    else { "Posible adaptador de red USB" }
                }}
            )

            # Candidatos seguros: USB con ambos criterios activos, sin ninguna exclusion.
            $excludedNames = @($excluded | ForEach-Object { $_.FriendlyName })
            $all = $all | Where-Object {
                $_.FriendlyName -match "USB" -and
                $_.Class        -match "USB" -and
                $_.FriendlyName -notin $excludedNames
            }
        }

        return @{
            Candidates = @($all | Sort-Object Class, FriendlyName | Select-Object Status, Class, FriendlyName, InstanceId)
            Excluded   = $excluded
        }
    }

    $driverList   = @($remoteData.Candidates)
    $excludedList = @($remoteData.Excluded)
    $total        = $driverList.Count

    # Mostrar Root Hubs excluidos (solo en modo todos — no afecta al modo fantasma)
    if (-not $soloFantasmas -and $excludedList.Count -gt 0) {
        Write-Warn "Excluidos $($excludedList.Count) dispositivo(s) USB (no seguros para borrar remotamente):"
        foreach ($e in $excludedList) {
            Append-Output ("    [EXCLUIDO]  {0,-26}  {1,-40}  ({2})" -f $e.Class, $e.FriendlyName, $e.ExcludeReason) ([System.Drawing.Color]::FromArgb(180, 80, 0))
        }
        Append-Output "" $script:White
    }

    if ($total -eq 0) {
        $noDriverMsg = if ($soloFantasmas) {
            "No se encontraron drivers USB fantasma en '$ComputerName'."
        } else {
            "No se encontraron drivers USB (excluidos Root Hubs) en '$ComputerName'."
        }
        Write-Warn $noDriverMsg
        Write-Sep
        return
    }

    Write-Info "Se encontraron $total driver(s) candidato(s) [$modoTexto]:"
    Append-Output "" $script:White

    $idx = 0
    foreach ($d in $driverList) {
        $idx++
        $stateColor = switch ($d.Status) {
            "OK"      { [System.Drawing.Color]::LightGreen  }
            "Error"   { [System.Drawing.Color]::Tomato      }
            "Unknown" { [System.Drawing.Color]::Gray        }
            default   { [System.Drawing.Color]::LightYellow }
        }
        Append-Output ("    [{0:D2}] {1,-10}  {2,-26}  {3}" -f $idx, $d.Status, $d.Class, $d.FriendlyName) $stateColor
        Append-Output ("           InstanceId: {0}" -f $d.InstanceId) ([System.Drawing.Color]::FromArgb(120, 120, 120))
    }
    Append-Output "" $script:White

    # ── Fase C: Confirmacion ─────────────────────────────────────────
    $confirmMsg = "Modo: $modoTexto`n`n" +
                  "Se van a eliminar $total driver(s) USB en '$ComputerName'.`n`n" +
                  "Esta operacion puede requerir reinicio y no es facilmente reversible.`n`n" +
                  "Confirmar borrado?"
    if (-not (Confirm-Action $confirmMsg)) {
        Write-Warn "Borrado cancelado por el usuario. No se ha eliminado nada."
        return
    }

    # ── Fase D: Borrado con progreso visible y timeout por driver ────
    Write-Sep
    Write-Info "Iniciando borrado de $total driver(s) [$modoTexto]..."
    Append-Output "" $script:White

    # Tiempo maximo de espera por driver antes de declarar TIMEOUT y continuar.
    # pnputil sobre un driver del stack USB puede quedar esperando liberaciones
    # que nunca llegan; sin timeout el script se bloquea indefinidamente.
    $timeoutSec = 15

    $okCount    = 0
    $warnCount  = 0
    $errorCount = 0
    $results    = [System.Collections.Generic.List[object]]::new()
    $idx        = 0

    foreach ($d in $driverList) {
        $idx++
        Write-Info ("  [{0}/{1}] Procesando: {2}" -f $idx, $total, $d.FriendlyName)
        # Fuerza repintado sincrono antes del Start-Job para que el usuario vea
        # el mensaje "Procesando" antes de que el UI thread entre en el bucle de espera.
        $script:outputBox.Update()

        # Lanzar el borrado en un job independiente.
        # Invoke-Command directo no tiene timeout de ejecucion: si pnputil espera la
        # liberacion de un dispositivo bloqueado del stack USB, congela el hilo UI
        # indefinidamente. Start-Job desacopla la ejecucion; Stop-Job la cancela.
        $job = Start-Job -ArgumentList $ComputerName, ([string]$d.InstanceId) -ScriptBlock {
            param($computer, $instanceId)
            try {
                $res = Invoke-Command -ComputerName $computer -ArgumentList $instanceId `
                    -ErrorAction SilentlyContinue -ScriptBlock {
                        param($id)
                        $out = pnputil /remove-device "$id" 2>&1
                        return @{ Output = ($out -join ' ').Trim(); ExitCode = $LASTEXITCODE }
                    }
                return $res
            } catch {
                return @{ Output = $_.Exception.Message; ExitCode = -1 }
            }
        }

        # Esperar con deadline manteniendo la GUI viva mediante DoEvents.
        # Intervalo de 200ms: suficientemente corto para no congelar la ventana.
        $deadline = (Get-Date).AddSeconds($timeoutSec)
        while ($job.State -eq 'Running' -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 200
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($job.State -eq 'Running') {
            # El job sigue activo al superar el deadline: cancelar y continuar.
            Stop-Job  $job
            Remove-Job $job -Force
            Write-Fail ("  [{0}/{1}] TIMEOUT -> {2}  (>{3}s, driver en uso o stack USB bloqueado)" -f `
                $idx, $total, $d.FriendlyName, $timeoutSec)
            $errorCount++
            $results.Add([PSCustomObject]@{ Status="ERROR"; Name=$d.FriendlyName; Detail="TIMEOUT (>${timeoutSec}s)" })
            continue
        }

        # Job completado: recoger resultado y limpiar
        $rem = Receive-Job $job
        Remove-Job $job -Force

        if ($null -eq $rem) {
            Write-Fail ("  [{0}/{1}] ERROR  -> {2}  (sin respuesta remota)" -f $idx, $total, $d.FriendlyName)
            $errorCount++
            $results.Add([PSCustomObject]@{ Status="ERROR"; Name=$d.FriendlyName; Detail="Sin respuesta remota" })
            continue
        }

        switch ($rem.ExitCode) {
            0 {
                Write-Ok   ("  [{0}/{1}] OK     -> {2}" -f $idx, $total, $d.FriendlyName)
                $okCount++
                $results.Add([PSCustomObject]@{ Status="OK";   Name=$d.FriendlyName; Detail="OK" })
            }
            3010 {
                # pnputil: exito pero se requiere reinicio para completar
                Write-Warn ("  [{0}/{1}] WARN   -> {2}  (requiere reinicio)" -f $idx, $total, $d.FriendlyName)
                $warnCount++
                $results.Add([PSCustomObject]@{ Status="WARN"; Name=$d.FriendlyName; Detail="Eliminado, pendiente reinicio" })
            }
            default {
                Write-Fail ("  [{0}/{1}] ERROR  -> {2}  | ExitCode={3} | {4}" -f `
                    $idx, $total, $d.FriendlyName, $rem.ExitCode, $rem.Output)
                $errorCount++
                $results.Add([PSCustomObject]@{ Status="ERROR"; Name=$d.FriendlyName; Detail="ExitCode=$($rem.ExitCode)" })
            }
        }
    }

    # ── Fase E: Resumen final ────────────────────────────────────────
    Append-Output "" $script:White
    Write-Sep
    Write-Info "RESUMEN DE BORRADO USB"
    Write-Sep
    Append-Output ("  Modo               : {0}" -f $modoTexto)    ([System.Drawing.Color]::Cyan)
    Append-Output ("  Total candidatos   : {0}" -f $total)        $script:White
    Append-Output ("  Eliminados OK      : {0}" -f $okCount)      ([System.Drawing.Color]::LightGreen)
    if ($warnCount -gt 0) {
        Append-Output ("  Requieren reinicio : {0}" -f $warnCount) ([System.Drawing.Color]::Yellow)
    }
    if ($errorCount -gt 0) {
        Append-Output ("  Fallidos           : {0}" -f $errorCount) ([System.Drawing.Color]::Tomato)
        Append-Output "" $script:White
        Write-Fail "  Drivers con error:"
        foreach ($r in ($results | Where-Object { $_.Status -eq "ERROR" })) {
            Append-Output ("    -> {0}  ({1})" -f $r.Name, $r.Detail) ([System.Drawing.Color]::Tomato)
        }
    }
    Write-Sep
    Append-Output "" $script:White

    # Ofrecer reinicio si hubo al menos un borrado exitoso (OK o pendiente)
    $deleted = $okCount + $warnCount
    if ($deleted -gt 0) {
        $reinicioMsg = "Se procesaron correctamente $deleted de $total driver(s)."
        if ($warnCount -gt 0) { $reinicioMsg += "`n$warnCount requieren reinicio para completarse." }
        $reinicioMsg += "`n`nReiniciar '$ComputerName' ahora?"
        if (Confirm-Action $reinicioMsg) {
            Restart-Computer -ComputerName $ComputerName -Force
            Write-Ok "Reiniciando '$ComputerName'..."
        }
    }
}

#endregion

#region ═══════════════════════════════════════════════════════════
# OPCION 5 - MANTENIMIENTO DEL SISTEMA (DISM / SFC / CHKDSK)
#═══════════════════════════════════════════════════════════════════

function Invoke-RemoteRepair {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Sep
    Write-Info "Reparacion del sistema en '$ComputerName'"
    Write-Sep
    Write-Info "  Secuencia: [1/2] DISM /Online /Cleanup-Image /RestoreHealth"
    Write-Info "             [2/2] sfc /scannow"
    Write-Info "  Tiempo estimado total: 15-45 minutos."
    Append-Output "" $script:White

    if (-not (Confirm-Action (
        "Se va a ejecutar en '$ComputerName':" + "`n`n" +
        "  [1/2] DISM /Online /Cleanup-Image /RestoreHealth" + "`n" +
        "  [2/2] sfc /scannow" + "`n`n" +
        "  - Tiempo estimado: 15-45 minutos." + "`n" +
        "  - DISM repara la imagen del sistema (WinSxS)." + "`n" +
        "  - SFC repara archivos protegidos del sistema operativo." + "`n" +
        "  - No requiere reinicio (salvo que SFC encuentre archivos en uso)." + "`n`n" +
        "Continuar?"
    ))) {
        Write-Warn "Operacion cancelada por el usuario."
        return
    }

    # ── FASE 1: DISM ─────────────────────────────────────────────────────────────
    Write-Sep
    Write-Info "[1/2] Iniciando DISM /Online /Cleanup-Image /RestoreHealth..."
    Set-Progress 0 "[1/2] DISM en ejecucion..."
    $script:outputBox.Update()

    $jobDism = Start-Job -ArgumentList $ComputerName -ScriptBlock {
        param($computer)
        try {
            $res = Invoke-Command -ComputerName $computer -ErrorAction SilentlyContinue -ScriptBlock {
                $out = dism /Online /Cleanup-Image /RestoreHealth 2>&1
                return @{ Output = ($out -join "`n").Trim(); ExitCode = $LASTEXITCODE }
            }
            return $res
        } catch {
            return @{ Output = $_.Exception.Message; ExitCode = -1 }
        }
    }

    $okDism = Wait-JobWithEvents $jobDism -TimeoutMinutes 60 -ProgressLabel "DISM" `
                                          -ProgressFrom 0 -ProgressTo 48
    if (-not $okDism) {
        Write-Fail "[1/2] DISM TIMEOUT (>60 min). Puede seguir ejecutandose en '$ComputerName'."
        Write-Warn "  Comprueba el log en: C:\Windows\Logs\DISM\dism.log"
        Set-Progress 0 "DISM timeout"
        Write-Sep; Append-Output "" $script:White
        return
    }

    $remDism = Receive-Job $jobDism; Remove-Job $jobDism -Force

    if ($null -eq $remDism) {
        Write-Fail "[1/2] Sin respuesta remota (DISM). Verifica WinRM con '$ComputerName'."
        Set-Progress 0 "Sin respuesta"
        Write-Sep; Append-Output "" $script:White
        return
    }

    # Mostrar salida DISM
    Write-Info "[1/2] Salida de DISM:"
    foreach ($line in ($remDism.Output -split "`n" | Where-Object { $_.Trim() -ne "" })) {
        $l      = $line.Trim()
        $lColor = if     ($l -match "Error|fallo|FAIL")               { [System.Drawing.Color]::Tomato      }
                  elseif ($l -match "Warning|Advertencia")            { [System.Drawing.Color]::Yellow      }
                  elseif ($l -match "completado|correctamente|100\.0%") { [System.Drawing.Color]::LightGreen }
                  else                                                 { $script:Silver                      }
        Append-Output "    $l" $lColor
    }
    if ($remDism.ExitCode -eq 0) {
        Write-Ok   "[1/2] DISM completado. ExitCode=0."
    } else {
        Write-Fail "[1/2] DISM ExitCode=$($remDism.ExitCode). Revisa: C:\Windows\Logs\DISM\dism.log"
    }
    Set-Progress 50 "[1/2] DISM completado - Iniciando SFC..."
    Append-Output "" $script:White

    # ── FASE 2: SFC ──────────────────────────────────────────────────────────────
    Write-Sep
    Write-Info "[2/2] Iniciando sfc /scannow..."
    Set-Progress 50 "[2/2] SFC en ejecucion..."
    $script:outputBox.Update()

    $jobSfc = Start-Job -ArgumentList $ComputerName -ScriptBlock {
        param($computer)
        try {
            $res = Invoke-Command -ComputerName $computer -ErrorAction SilentlyContinue -ScriptBlock {
                # sfc puede producir salida UTF-16 ilegible via WinRM; se limpia despues.
                # El log CBS es la fuente fiable del resultado real.
                $out = sfc /scannow 2>&1
                $cbs = Get-Content "$env:windir\Logs\CBS\CBS.log" -Tail 20 -ErrorAction SilentlyContinue
                return @{
                    Output   = ($out -join "`n").Trim()
                    CbsTail  = if ($cbs) { ($cbs -join "`n").Trim() } else { "" }
                    ExitCode = $LASTEXITCODE
                }
            }
            return $res
        } catch {
            return @{ Output = $_.Exception.Message; CbsTail = ""; ExitCode = -1 }
        }
    }

    $okSfc = Wait-JobWithEvents $jobSfc -TimeoutMinutes 30 -ProgressLabel "SFC" `
                                        -ProgressFrom 50 -ProgressTo 98
    if (-not $okSfc) {
        Write-Fail "[2/2] SFC TIMEOUT (>30 min). Puede seguir ejecutandose en '$ComputerName'."
        Set-Progress 50 "SFC timeout"
        Write-Sep; Append-Output "" $script:White
        return
    }

    $remSfc = Receive-Job $jobSfc; Remove-Job $jobSfc -Force

    if ($null -eq $remSfc) {
        Write-Fail "[2/2] Sin respuesta remota (SFC). Verifica WinRM con '$ComputerName'."
        Set-Progress 50 "Sin respuesta SFC"
        Write-Sep; Append-Output "" $script:White
        return
    }

    # Mostrar salida SFC (limpiar artefactos UTF-16) y log CBS
    Write-Info "[2/2] Salida SFC:"
    $sfcOut = ($remSfc.Output -replace "[^\x20-\x7E\r\n]", "").Trim()
    if ($sfcOut) {
        foreach ($line in ($sfcOut -split "`n" | Where-Object { $_.Trim() } | Select-Object -Last 8)) {
            Append-Output "    $($line.Trim())" $script:Silver
        }
    }
    if ($remSfc.CbsTail) {
        Write-Info "[2/2] Tail del log CBS:"
        foreach ($line in ($remSfc.CbsTail -split "`n" | Where-Object { $_.Trim() })) {
            $l      = $line.Trim()
            $lColor = if     ($l -match "error|fail|corrupt")    { [System.Drawing.Color]::Tomato      }
                      elseif ($l -match "repaired|fixed|reparo")  { [System.Drawing.Color]::LightGreen }
                      else                                        { $script:Silver                      }
            Append-Output "    $l" $lColor
        }
    }
    switch ($remSfc.ExitCode) {
        0 { Write-Ok   "[2/2] SFC completado. ExitCode=0." }
        1 { Write-Warn "[2/2] SFC: archivos danados no reparados. ExitCode=1." }
        2 { Write-Ok   "[2/2] SFC: limpieza realizada. ExitCode=2." }
        3 { Write-Warn "[2/2] SFC no completo el escaneo. Puede requerir reinicio. ExitCode=3." }
        default { Write-Fail "[2/2] SFC ExitCode=$($remSfc.ExitCode). Revisa: C:\Windows\Logs\CBS\CBS.log" }
    }

    # ── Resumen ───────────────────────────────────────────────────────────────────
    Set-Progress 100 "Reparacion completada"
    Write-Sep
    Write-Info "Resumen reparacion '$ComputerName':"
    if ($remDism.ExitCode -eq 0) { Write-Ok   "  DISM : Completado (ExitCode=0)" }
    else                          { Write-Warn "  DISM : Con avisos  (ExitCode=$($remDism.ExitCode))" }
    $sfcOk = $remSfc.ExitCode -eq 0 -or $remSfc.ExitCode -eq 2
    if ($sfcOk) { Write-Ok   "  SFC  : Completado (ExitCode=$($remSfc.ExitCode))" }
    else         { Write-Warn "  SFC  : Con avisos  (ExitCode=$($remSfc.ExitCode))" }
    Append-Output "" $script:White
}


function Invoke-RemoteChkdsk {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Sep
    Write-Info "CHKDSK /r en '$ComputerName'"
    Write-Sep
    Write-Info "  En la unidad del sistema (C:) normalmente requiere reinicio para ejecutarse."
    Write-Info "  En otras unidades puede ejecutarse en caliente si no hay archivos bloqueados."
    Append-Output "" $script:White

    $driveLetter = Get-Input "Letra de unidad a verificar (solo la letra, p.ej. C)" "ChkDsk - Unidad" "C"
    if ([string]::IsNullOrWhiteSpace($driveLetter)) {
        Write-Warn "Operacion cancelada por el usuario."
        return
    }
    $driveLetter = $driveLetter.Trim().TrimEnd(":").ToUpper()
    if ($driveLetter -notmatch "^[A-Z]$") {
        Write-Fail "Letra de unidad no valida: '$driveLetter'. Debe ser una sola letra (A-Z)."
        return
    }

    $driveTarget   = "${driveLetter}:"
    $isSystemDrive = ($driveLetter -eq "C")
    $extraWarning  = if ($isSystemDrive) {
        "`n`n  IMPORTANTE: '$driveTarget' es probablemente la unidad del sistema." +
        "`n  ChkDsk no puede bloquearla mientras Windows esta en ejecucion." +
        "`n  Quedara PROGRAMADO para ejecutarse en el siguiente reinicio."
    } else { "" }

    if (-not (Confirm-Action (
        "Se va a ejecutar en '$ComputerName':" + "`n`n" +
        "  chkdsk $driveTarget /r" + "`n`n" +
        "  - /r localiza sectores defectuosos y recupera informacion legible." + "`n" +
        "  - En el volumen del sistema requiere reinicio para ejecutarse." + "`n" +
        "  - En otras unidades puede tardar varios minutos." +
        $extraWarning + "`n`n" +
        "Continuar?"
    ))) {
        Write-Warn "Operacion cancelada por el usuario."
        return
    }

    Write-Info "Ejecutando chkdsk $driveTarget /r en '$ComputerName'..."
    $script:outputBox.Update()

    $job = Start-Job -ArgumentList $ComputerName, $driveTarget -ScriptBlock {
        param($computer, $drive)
        try {
            $res = Invoke-Command -ComputerName $computer -ArgumentList $drive -ErrorAction SilentlyContinue -ScriptBlock {
                param($d)
                # En la unidad del sistema chkdsk responde de inmediato indicando que se
                # programara para el siguiente reinicio (no bloquea). En otras unidades
                # puede tardar (scan de sectores fisicos). /r incluye /f implicito.
                $out = chkdsk $d /r 2>&1
                return @{ Output = ($out -join "`n").Trim(); ExitCode = $LASTEXITCODE }
            }
            return $res
        } catch {
            return @{ Output = $_.Exception.Message; ExitCode = -1 }
        }
    }

    $ok = Wait-JobWithEvents $job -TimeoutMinutes 60 -ProgressLabel "ChkDsk $driveTarget"
    if (-not $ok) {
        Write-Fail "ChkDsk TIMEOUT (>60 min). La operacion puede seguir en curso en '$ComputerName'."
        Write-Sep; Append-Output "" $script:White
        return
    }

    $rem = Receive-Job $job; Remove-Job $job -Force

    if ($null -eq $rem) {
        Write-Fail "Sin respuesta remota. Verifica la conectividad WinRM con '$ComputerName'."
        Write-Sep; Append-Output "" $script:White
        return
    }

    Write-Sep
    Write-Info "Salida de ChkDsk ${driveTarget}:"
    foreach ($line in ($rem.Output -split "`n" | Where-Object { $_.Trim() })) {
        $l      = $line.Trim()
        $lColor = if     ($l -match "programar|reinicio|reboot|schedule|restart|siguiente")  { [System.Drawing.Color]::Yellow      }
                  elseif ($l -match "error|danado|corrupt|bad sector")                        { [System.Drawing.Color]::Tomato      }
                  elseif ($l -match "correcto|completado|sin errores|no errors|Windows comprobado") { [System.Drawing.Color]::LightGreen }
                  else                                                                         { $script:Silver                      }
        Append-Output "    $l" $lColor
    }
    Write-Sep

    # Detectar caso "programado para reinicio" por contenido de salida.
    # ExitCode puede ser 0 en este caso, por eso se comprueba la salida primero.
    $scheduledForReboot = ($rem.Output -match "programar|schedule|reinicio|siguiente arranque|next restart|next boot")
    if ($scheduledForReboot) {
        Write-Warn "ChkDsk $driveTarget queda PROGRAMADO para el siguiente reinicio de '$ComputerName'."
        Write-Warn "  No pudo bloquear el volumen en caliente (comportamiento normal en unidad del sistema)."
        Write-Info "  ChkDsk se ejecutara automaticamente al arrancar el equipo."
        Append-Output "" $script:White
        if (Confirm-Action "ChkDsk $driveTarget esta programado para el reinicio.`n`nReiniciar '$ComputerName' ahora para ejecutarlo?") {
            try {
                Restart-Computer -ComputerName $ComputerName -Force -ErrorAction Stop
                Write-Ok "Orden de reinicio enviada a '$ComputerName'."
            } catch {
                Write-Fail "No se pudo reiniciar '$ComputerName': $_"
            }
        } else {
            Write-Info "  Recuerda reiniciar '$ComputerName' para que ChkDsk se ejecute."
        }
    } elseif ($rem.ExitCode -eq 0) {
        Write-Ok "ChkDsk $driveTarget completado correctamente. ExitCode=0."
    } elseif ($rem.ExitCode -eq 1) {
        Write-Warn "ChkDsk encontro y corrigio errores en $driveTarget. ExitCode=1."
        Write-Info "  Puede ser recomendable reiniciar para completar las reparaciones."
    } elseif ($rem.ExitCode -eq 2) {
        Write-Warn "ChkDsk realizo limpieza en $driveTarget. ExitCode=2."
    } else {
        Write-Fail "ChkDsk $driveTarget finalizo con ExitCode=$($rem.ExitCode)."
    }
    Append-Output "" $script:White
}

#endregion

#region ═══════════════════════════════════════════════════════════
# INTERFAZ GRAFICA (WinForms)
#═══════════════════════════════════════════════════════════════════

# ── Colores y fuentes ────────────────────────────────────────────
$bgDark      = [System.Drawing.Color]::FromArgb(28,  28,  28)
$bgPanel     = [System.Drawing.Color]::FromArgb(45,  45,  48)
$bgOutput    = [System.Drawing.Color]::FromArgb(16,  16,  16)
$accent      = [System.Drawing.Color]::FromArgb(0,   122, 204)
$btnRed      = [System.Drawing.Color]::FromArgb(160,  40,  40)
$btnGray     = [System.Drawing.Color]::FromArgb(62,   62,  66)
$white       = $script:White
$silver      = $script:Silver

$fontUI    = New-Object System.Drawing.Font("Segoe UI",  10)
$fontMono  = New-Object System.Drawing.Font("Consolas",   9)
$fontTitle = New-Object System.Drawing.Font("Segoe UI",  13, [System.Drawing.FontStyle]::Bold)
$fontSmall = New-Object System.Drawing.Font("Segoe UI",   9)

# ── Formulario ───────────────────────────────────────────────────
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Administracion Remota  |  Accenture / Airbus"
$form.Size            = New-Object System.Drawing.Size(1130, 680)
$form.MinimumSize     = New-Object System.Drawing.Size(920, 500)
$form.BackColor       = $bgDark
$form.ForeColor       = $white
$form.Font            = $fontUI
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "Sizable"

# ── Helper para crear botones planos ─────────────────────────────
function New-FlatButton {
    param(
        [string]$Text,
        [int]$X, [int]$Y,
        [int]$Width  = 195,
        [int]$Height = 30,
        [System.Drawing.Color]$Bg
    )
    $b                            = New-Object System.Windows.Forms.Button
    $b.Text                       = $Text
    $b.Location                   = New-Object System.Drawing.Point($X, $Y)
    $b.Size                       = New-Object System.Drawing.Size($Width, $Height)
    $b.BackColor                  = $Bg
    $b.ForeColor                  = $white
    $b.FlatStyle                  = "Flat"
    $b.FlatAppearance.BorderSize  = 1
    $b.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $b.Cursor                     = "Hand"
    $b.Font                       = $fontSmall
    return $b
}

# ── Panel superior ────────────────────────────────────────────────
$topPanel           = New-Object System.Windows.Forms.Panel
$topPanel.Dock      = "Top"
$topPanel.Height    = 215
$topPanel.BackColor = $bgPanel
$form.Controls.Add($topPanel)

# Titulo y subtitulo
$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "  ADMINISTRACION REMOTA"
$lblTitle.Font      = $fontTitle
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 190, 255)
$lblTitle.AutoSize  = $true
$lblTitle.Location  = New-Object System.Drawing.Point(8, 8)
$topPanel.Controls.Add($lblTitle)

$lblSub             = New-Object System.Windows.Forms.Label
$lblSub.Text        = "  Accenture / Airbus  |  PowerShell 5.1"
$lblSub.Font        = $fontSmall
$lblSub.ForeColor   = $silver
$lblSub.AutoSize    = $true
$lblSub.Location    = New-Object System.Drawing.Point(8, 34)
$topPanel.Controls.Add($lblSub)

# Campo equipo
$lblEquipo          = New-Object System.Windows.Forms.Label
$lblEquipo.Text     = "Equipo:"
$lblEquipo.ForeColor= $silver
$lblEquipo.AutoSize = $true
$lblEquipo.Location = New-Object System.Drawing.Point(12, 60)
$topPanel.Controls.Add($lblEquipo)

$txtEquipo             = New-Object System.Windows.Forms.TextBox
$txtEquipo.Width       = 260
$txtEquipo.Location    = New-Object System.Drawing.Point(68, 57)
$txtEquipo.BackColor   = [System.Drawing.Color]::FromArgb(55, 55, 58)
$txtEquipo.ForeColor   = $white
$txtEquipo.BorderStyle = "FixedSingle"
$txtEquipo.Font        = $fontUI
$topPanel.Controls.Add($txtEquipo)

$btnPing = New-FlatButton "Ping" 340 57 60 30 $btnGray
$topPanel.Controls.Add($btnPing)

# Label de resultado de ping: ONLINE|VPN|IP / ONLINE|CABLE|IP / OFFLINE
$lblPingResult             = New-Object System.Windows.Forms.Label
$lblPingResult.Text        = ""
$lblPingResult.Location    = New-Object System.Drawing.Point(408, 62)
$lblPingResult.Size        = New-Object System.Drawing.Size(490, 22)
$lblPingResult.ForeColor   = $white
$lblPingResult.BackColor   = [System.Drawing.Color]::Transparent
$lblPingResult.Font        = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$topPanel.Controls.Add($lblPingResult)

# Referencia compartida para que los closures del panel lateral accedan al textbox
$script:EquipoInputBox = $txtEquipo

# Selector de modo Nacional / Divisional
$lblModo           = New-Object System.Windows.Forms.Label
$lblModo.Text      = "Modo:"
$lblModo.ForeColor = $silver
$lblModo.AutoSize  = $true
$lblModo.Location  = New-Object System.Drawing.Point(12, 93)
$topPanel.Controls.Add($lblModo)

$rbNacional             = New-Object System.Windows.Forms.RadioButton
$rbNacional.Text        = "Nacional"
$rbNacional.ForeColor   = $white
$rbNacional.AutoSize    = $true
$rbNacional.Checked     = $true
$rbNacional.Location    = New-Object System.Drawing.Point(58, 90)
$rbNacional.Font        = $fontSmall
$topPanel.Controls.Add($rbNacional)

$rbDivisional           = New-Object System.Windows.Forms.RadioButton
$rbDivisional.Text      = "Divisional"
$rbDivisional.ForeColor = $white
$rbDivisional.AutoSize  = $true
$rbDivisional.Location  = New-Object System.Drawing.Point(145, 90)
$rbDivisional.Font      = $fontSmall
$topPanel.Controls.Add($rbDivisional)

$rbNacional.Add_CheckedChanged({
    $script:Modo = if ($rbNacional.Checked) { "Nacional" } else { "Divisional" }
})

# Separador visual
$sep           = New-Object System.Windows.Forms.Panel
$sep.Location  = New-Object System.Drawing.Point(0, 114)
$sep.Size      = New-Object System.Drawing.Size(2000, 1)
$sep.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$topPanel.Controls.Add($sep)

# ── Fila 1: acciones principales ─────────────────────────────────
$btnMaster   = New-FlatButton "  Masterizacion"     10 120 175 28 $accent
$btnSoftware = New-FlatButton "  Software SCCM"    190 120 175 28 $accent
$btnInfo     = New-FlatButton "  Info del Sistema" 370 120 150 28 ([System.Drawing.Color]::FromArgb(40, 110, 60))
$btnUsb      = New-FlatButton "  Borrar USB"       525 120 130 28 $btnRed
$btnClear    = New-FlatButton "  Limpiar"          670 120  75 28 $btnGray
$btnCancel   = New-FlatButton "  Cancelar"         750 120  90 28 ([System.Drawing.Color]::FromArgb(200, 100, 0))

# ── Fila 2: reinicio (boton visible) + desplegable de mantenimiento ──
$btnRestart = New-FlatButton "  Reiniciar" 10 153 140 28 ([System.Drawing.Color]::FromArgb(160, 80, 0))

$cboMaintenance                  = New-Object System.Windows.Forms.ComboBox
$cboMaintenance.DropDownStyle    = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboMaintenance.BackColor        = [System.Drawing.Color]::FromArgb(55, 55, 58)
$cboMaintenance.ForeColor        = $white
$cboMaintenance.Font             = $fontSmall
$cboMaintenance.Location         = New-Object System.Drawing.Point(155, 156)
$cboMaintenance.Size             = New-Object System.Drawing.Size(385, 26)
$cboMaintenance.Items.AddRange(@(
    "GPUpdate /force",
    "Ciclos SCCM",
    "Reparacion sistema (DISM + SFC)",
    "ChkDsk /r"
))
$cboMaintenance.SelectedIndex = 0
$topPanel.Controls.Add($cboMaintenance)

$btnExecute = New-FlatButton "  Ejecutar" 545 153 110 28 ([System.Drawing.Color]::FromArgb(0, 130, 60))

# ── Fila 3: barra de progreso (DISM + SFC) ────────────────────────
$script:progressBar          = New-Object System.Windows.Forms.ProgressBar
$script:progressBar.Location = New-Object System.Drawing.Point(10, 186)
$script:progressBar.Size     = New-Object System.Drawing.Size(535, 18)
$script:progressBar.Minimum  = 0
$script:progressBar.Maximum  = 100
$script:progressBar.Value    = 0
$script:progressBar.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
$topPanel.Controls.Add($script:progressBar)

$script:progressLabel           = New-Object System.Windows.Forms.Label
$script:progressLabel.Location  = New-Object System.Drawing.Point(552, 186)
$script:progressLabel.Size      = New-Object System.Drawing.Size(260, 18)
$script:progressLabel.ForeColor = $silver
$script:progressLabel.Font      = $fontSmall
$script:progressLabel.Text      = ""
$script:progressLabel.TextAlign = "MiddleLeft"
$topPanel.Controls.Add($script:progressLabel)

$btnCancel.Enabled = $false

foreach ($b in @($btnMaster,$btnSoftware,$btnInfo,$btnUsb,$btnClear,$btnCancel,$btnRestart,$btnExecute)) {
    $topPanel.Controls.Add($b)
}

# ── Area de salida ────────────────────────────────────────────────
$script:outputBox             = New-Object System.Windows.Forms.RichTextBox
$script:outputBox.Dock        = "Fill"
$script:outputBox.BackColor   = $bgOutput
$script:outputBox.ForeColor   = $white
$script:outputBox.Font        = $fontMono
$script:outputBox.ReadOnly    = $true
$script:outputBox.BorderStyle = "None"
$script:outputBox.ScrollBars  = "Vertical"
$script:outputBox.WordWrap    = $false
$form.Controls.Add($script:outputBox)

# ── Barra de estado ───────────────────────────────────────────────
$statusBar          = New-Object System.Windows.Forms.Panel
$statusBar.Dock     = "Bottom"
$statusBar.Height   = 24
$statusBar.BackColor= $accent
$form.Controls.Add($statusBar)

# ── Panel lateral derecho (equipos en seguimiento) ────────────────
# Se anade ANTES de BringToFront para que el Dock=Right se resuelva
# antes que el Dock=Fill del outputBox, dejando la columna derecha libre.
$rightPanel             = New-Object System.Windows.Forms.Panel
$rightPanel.Dock        = "Right"
$rightPanel.Width       = 210
$rightPanel.BackColor   = [System.Drawing.Color]::FromArgb(38, 38, 42)
$form.Controls.Add($rightPanel)

$rightHeader            = New-Object System.Windows.Forms.Label
$rightHeader.Text       = "Equipos masterizando"
$rightHeader.Dock       = "Top"
$rightHeader.Height     = 32
$rightHeader.Font       = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$rightHeader.ForeColor  = [System.Drawing.Color]::FromArgb(0, 190, 255)
$rightHeader.BackColor  = [System.Drawing.Color]::FromArgb(28, 28, 28)
$rightHeader.TextAlign  = "MiddleCenter"
$rightPanel.Controls.Add($rightHeader)

$rightBtnPanel          = New-Object System.Windows.Forms.Panel
$rightBtnPanel.Dock     = "Bottom"
$rightBtnPanel.Height   = 96
$rightBtnPanel.BackColor= [System.Drawing.Color]::FromArgb(28, 28, 28)
$rightPanel.Controls.Add($rightBtnPanel)

$btnAddEquipo      = New-FlatButton "Anadir equipo"      3  3 202 26 ([System.Drawing.Color]::FromArgb(0, 100, 50))
$btnRemoveEquipo   = New-FlatButton "Quitar selec."      3 33 202 26 $btnGray
$btnRefreshEquipos = New-FlatButton "Refrescar estado"   3 63 202 26 ([System.Drawing.Color]::FromArgb(30, 60, 110))
$rightBtnPanel.Controls.Add($btnAddEquipo)
$rightBtnPanel.Controls.Add($btnRemoveEquipo)
$rightBtnPanel.Controls.Add($btnRefreshEquipos)

$script:flowEquipos               = New-Object System.Windows.Forms.FlowLayoutPanel
$script:flowEquipos.Dock          = "Fill"
$script:flowEquipos.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$script:flowEquipos.AutoScroll    = $true
$script:flowEquipos.WrapContents  = $true
$script:flowEquipos.BackColor     = [System.Drawing.Color]::FromArgb(32, 32, 35)
$script:flowEquipos.Padding       = New-Object System.Windows.Forms.Padding(3, 8, 0, 3)
$rightPanel.Controls.Add($script:flowEquipos)

$script:outputBox.BringToFront()   # Fill solo ocupa el espacio restante

$script:statusLabel           = New-Object System.Windows.Forms.Label
$script:statusLabel.Text      = "  Listo - introduce un equipo y pulsa una accion"
$script:statusLabel.Dock      = "Fill"
$script:statusLabel.ForeColor = $white
$script:statusLabel.Font      = $fontSmall
$script:statusLabel.TextAlign = "MiddleLeft"
$statusBar.Controls.Add($script:statusLabel)

# ── Helpers del panel lateral de equipos ─────────────────────────

# Archivo de persistencia en el mismo directorio que el script
$script:EquiposFile = Join-Path $PSScriptRoot "equipos_seguimiento.json"

# ── Click handler compartido ──────────────────────────────────────
# Usa [System.EventHandler] para recibir el sender directamente, eliminando
# los problemas de closure de PS 5.1 con variables de funcion capturadas.
# Sube por la jerarquia de controles hasta encontrar el Panel-tarjeta
# (reconocible porque su Tag contiene el nombre del equipo).
$script:CardClickHandler = [System.EventHandler]{
    param($sender, $e)
    $ctrl = $sender
    while ($ctrl -and -not ($ctrl -is [System.Windows.Forms.Panel] -and $ctrl.Tag)) {
        $ctrl = $ctrl.Parent
    }
    if ($ctrl -and $ctrl.Tag) {
        $script:EquipoInputBox.Text = $ctrl.Tag
        Set-EquipoSeleccionado $ctrl
    }
}

function Set-EquipoSeleccionado {
    param($CardPanel)
    foreach ($c in $script:flowEquipos.Controls) {
        if ($c -is [System.Windows.Forms.Panel]) {
            $c.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 58)
        }
    }
    if ($CardPanel) { $CardPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 80, 140) }
    $script:EquipoSelCard = $CardPanel
}

function New-EquipoCard {
    param([string]$Name)

    $card           = New-Object System.Windows.Forms.Panel
    $card.Tag       = $Name
    $card.Width     = 175
    $card.Height    = 52
    $card.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 58)
    $card.Cursor    = "Hand"
    $card.Margin    = New-Object System.Windows.Forms.Padding(5, 5, 5, 5)

    # Un unico Label por tarjeta para evitar nombre duplicado y el problema
    # de BackColor=Transparent en el primer control anadido durante Add_Shown.
    # Formato: "  ... | PCNAME" → "  ONLINE  |  PCNAME" / "  OFFLINE  |  PCNAME"
    $lStatus           = New-Object System.Windows.Forms.Label
    $lStatus.Name      = "lblStatus"
    $lStatus.Text      = "  ... | $Name"
    $lStatus.ForeColor = [System.Drawing.Color]::Gray
    $lStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lStatus.Location  = New-Object System.Drawing.Point(6, 16)
    $lStatus.Size      = New-Object System.Drawing.Size(163, 20)
    $lStatus.BackColor = [System.Drawing.Color]::Transparent
    $lStatus.Cursor    = "Hand"
    $card.Controls.Add($lStatus)

    $card.Add_Click($script:CardClickHandler)
    $lStatus.Add_Click($script:CardClickHandler)

    return $card
}

function Update-EquipoCard {
    param($CardPanel)
    $lStatus = $CardPanel.Controls | Where-Object { $_.Name -eq "lblStatus" }
    if (-not $lStatus) { return }
    $online = Test-Connection -ComputerName $CardPanel.Tag -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($online) {
        $lStatus.Text      = "  ONLINE  |  $($CardPanel.Tag)"
        $lStatus.ForeColor = [System.Drawing.Color]::LightGreen
    } else {
        $lStatus.Text      = "  OFFLINE  |  $($CardPanel.Tag)"
        $lStatus.ForeColor = [System.Drawing.Color]::Tomato
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Refresh-EquipoEstados {
    foreach ($card in @($script:flowEquipos.Controls)) {
        if ($card -is [System.Windows.Forms.Panel] -and $card.Tag) {
            Update-EquipoCard $card
        }
    }
}

# ── Persistencia de la lista de equipos ───────────────────────────

function Save-EquipoList {
    try {
        $names = @($script:flowEquipos.Controls |
            Where-Object { $_ -is [System.Windows.Forms.Panel] -and $_.Tag } |
            ForEach-Object { $_.Tag })
        # -InputObject pasa el array completo como un objeto unico → JSON de array valido.
        # Piping elemento a elemento produciria strings JSON separadas (malformado) y
        # con array vacio no escribiria nada, dejando el archivo con contenido antiguo.
        $json = if ($names.Count -gt 0) { ConvertTo-Json -InputObject $names -Compress } else { '[]' }
        Set-Content -Path $script:EquiposFile -Value $json -Encoding UTF8
    } catch { <# sin permisos de escritura: se ignora silenciosamente #> }
}

function Load-EquipoList {
    if (-not (Test-Path $script:EquiposFile)) { return }
    try {
        $raw  = Get-Content -Path $script:EquiposFile -Raw -Encoding UTF8
        $data = $raw | ConvertFrom-Json
        # ConvertFrom-Json devuelve string si hay 1 elemento, array si hay varios.
        # @() normaliza ambos casos.
        foreach ($name in @($data)) {
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $already = $script:flowEquipos.Controls |
                Where-Object { $_ -is [System.Windows.Forms.Panel] -and $_.Tag -eq $name }
            if ($already) { continue }
            $card = New-EquipoCard $name
            $script:flowEquipos.Controls.Add($card)
        }
    } catch { <# archivo corrupto: se ignora silenciosamente #> }
}

# ── Helpers de control de UI ──────────────────────────────────────

# Lista de botones de accion (todos excepto Cancelar y Limpiar)
$script:ActionButtons = @($btnMaster,$btnSoftware,$btnInfo,$btnUsb,$btnRestart,$btnExecute,$btnPing)

function Set-ButtonsEnabled {
    param([bool]$Enabled)
    foreach ($b in $script:ActionButtons) { $b.Enabled = $Enabled }
    $cboMaintenance.Enabled = $Enabled
    $btnCancel.Enabled      = -not $Enabled
}

function Get-ValidComputer {
    $computer = $txtEquipo.Text.Trim()
    if ([string]::IsNullOrEmpty($computer)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Introduce el nombre del equipo remoto.",
            "Campo requerido",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return $null
    }
    Set-Status "Comprobando conectividad con '$computer'..." ([System.Drawing.Color]::Yellow)
    if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
        Write-Fail "El equipo '$computer' no esta accesible."
        Set-Status "Equipo no accesible" ([System.Drawing.Color]::Tomato)
        return $null
    }
    Write-Ok "Equipo '$computer' online."
    return $computer
}

# Wrappea el ciclo de vida comun de un boton de accion:
# disable -> ejecutar accion -> enable -> actualizar status
function Invoke-ActionButton {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$StatusMsg,
        [Parameter(Mandatory)][scriptblock]$Action,
        [bool]$UseCancel = $true
    )
    if ($UseCancel) { $script:cancelRequested = $false }
    Set-ButtonsEnabled $false
    Set-Status $StatusMsg ([System.Drawing.Color]::Yellow)
    try   { & $Action }
    catch { Write-Fail "Error inesperado: $_" }
    Set-ButtonsEnabled $true
    if ($UseCancel) {
        $finalMsg   = if ($script:cancelRequested) { "Cancelado"  } else { "Finalizado"  }
        $finalColor = if ($script:cancelRequested) { [System.Drawing.Color]::Orange } else { [System.Drawing.Color]::LightGreen }
        Set-Status $finalMsg $finalColor
        $script:cancelRequested = $false
    }
}

# ── Eventos ───────────────────────────────────────────────────────

$btnPing.Add_Click({
    $computer = $txtEquipo.Text.Trim()
    if ([string]::IsNullOrEmpty($computer)) { return }

    $lblPingResult.Text      = "Comprobando..."
    $lblPingResult.ForeColor = [System.Drawing.Color]::Yellow
    Set-Status "Haciendo ping a '$computer'..." ([System.Drawing.Color]::Yellow)
    Write-Sep

    if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
        # Resolver IP y clasificar tipo de conexion
        $ip = $null
        try {
            $ip = [System.Net.Dns]::GetHostAddresses($computer) |
                  Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
                  Select-Object -First 1 -ExpandProperty IPAddressToString
        } catch { }

        $tipo  = if ($ip -and ($ip.StartsWith("10.142.") -or $ip.StartsWith("10.99."))) { "VPN" } else { "CABLE" }
        $ipStr = if ($ip) { $ip } else { "?" }

        $lblPingResult.Text      = "ONLINE  |  $tipo  |  $ipStr"
        $lblPingResult.ForeColor = [System.Drawing.Color]::LightGreen
        Set-Status "  '$computer' ONLINE  |  $tipo  |  $ipStr" ([System.Drawing.Color]::LightGreen)
        Write-Ok "Ping OK -> '$computer' | Tipo: $tipo | IP: $ipStr"
    } else {
        $lblPingResult.Text      = "OFFLINE"
        $lblPingResult.ForeColor = [System.Drawing.Color]::Tomato
        Set-Status "  '$computer' OFFLINE" ([System.Drawing.Color]::Tomato)
        Write-Fail "Ping FAIL -> '$computer' no responde."
    }
    Write-Sep
    Append-Output "" $white
})

$btnMaster.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target `
        -StatusMsg "Comprobando masterizacion de '$target'..." `
        -Action    { Invoke-MasterCheck -ComputerName $target }
})

$btnSoftware.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target `
        -StatusMsg "Consultando Centro de Software de '$target'..." `
        -Action    { Invoke-SoftwareCheck -ComputerName $target }
})

$btnInfo.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target `
        -StatusMsg  "Obteniendo informacion de '$target'..." `
        -Action     { Invoke-SystemInfo -ComputerName $target } `
        -UseCancel  $false
    Set-Status "Finalizado" ([System.Drawing.Color]::LightGreen)
})

$btnUsb.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target `
        -StatusMsg "Obteniendo drivers USB de '$target'..." `
        -Action    { Invoke-UsbDriverClean -ComputerName $target }
})

$btnRestart.Add_Click({
    $target = $txtEquipo.Text.Trim()
    if ([string]::IsNullOrEmpty($target)) { return }
    if (-not (Confirm-Action "ATENCION: Se va a reiniciar el equipo '$target'.`n`nTodas las sesiones y procesos activos se cerraran.`n`n¿Confirmas el reinicio?" "Confirmar reinicio remoto")) {
        Write-Warn "Reinicio cancelado por el usuario."
        return
    }
    Set-ButtonsEnabled $false
    Set-Status "Reiniciando '$target'..." ([System.Drawing.Color]::Orange)
    Write-Sep
    Write-Info "Reinicio remoto de '$target'..."
    try {
        Restart-Computer -ComputerName $target -Force -ErrorAction Stop
        Write-Ok "Orden de reinicio enviada a '$target'."
        Set-Status "Reinicio enviado a '$target'" ([System.Drawing.Color]::LightGreen)
    } catch {
        Write-Fail "No se pudo reiniciar '$target': $_"
        Set-Status "Error al reiniciar" ([System.Drawing.Color]::Tomato)
    }
    Write-Sep
    Append-Output "" $white
    Set-ButtonsEnabled $true
})

$btnExecute.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    $opcion = $cboMaintenance.SelectedItem
    switch ($opcion) {
        "GPUpdate /force" {
            Invoke-ActionButton -ComputerName $target -UseCancel $false `
                -StatusMsg "Ejecutando gpupdate /force en '$target'..." `
                -Action {
                    Write-Sep
                    Write-Info "gpupdate /force en '$target'..."
                    $res = Invoke-RemoteGpupdate -ComputerName $target
                    Write-Sep
                    Append-Output "" $script:White
                    if ($res -and $res.Status -eq "OK") {
                        Write-Ok "gpupdate completado correctamente ($($res.Details))."
                        Set-Status "GPUpdate OK en '$target'" ([System.Drawing.Color]::LightGreen)
                    } else {
                        $detail = if ($res) { $res.Details } else { "Sin respuesta remota" }
                        Write-Warn "gpupdate: $detail."
                        Set-Status "GPUpdate WARN en '$target'" ([System.Drawing.Color]::Yellow)
                    }
                }
        }
        "Ciclos SCCM" {
            Invoke-ActionButton -ComputerName $target -UseCancel $false `
                -StatusMsg "Lanzando ciclos SCCM en '$target'..." `
                -Action {
                    Write-Sep
                    Write-Info "Ciclos de politicas SCCM en '$target'..."
                    $result = Invoke-Command -ComputerName $target -ScriptBlock $script:SccmCyclesBlock
                    Write-Sep
                    Append-Output "" $script:White
                    if ($result) {
                        switch ($result.Status) {
                            "OK"    { Write-Ok   "Ciclos completados: $($result.Details)"
                                      Set-Status "Ciclos SCCM OK en '$target'" ([System.Drawing.Color]::LightGreen) }
                            "WARN"  { Write-Warn "Ciclos con avisos: $($result.Details)"
                                      Set-Status "Ciclos SCCM con avisos" ([System.Drawing.Color]::Yellow) }
                            "ERROR" { Write-Fail "Ciclos con errores: $($result.Details)"
                                      Set-Status "Error en ciclos SCCM" ([System.Drawing.Color]::Tomato) }
                        }
                    } else {
                        Write-Warn "Sin respuesta del cliente SCCM. Verifica que CcmExec esta activo."
                        Set-Status "Sin respuesta SCCM en '$target'" ([System.Drawing.Color]::Yellow)
                    }
                }
        }
        "Reparacion sistema (DISM + SFC)" {
            Set-Progress 0 ""
            Invoke-ActionButton -ComputerName $target -UseCancel $false `
                -StatusMsg "Reparacion del sistema en '$target'..." `
                -Action    { Invoke-RemoteRepair -ComputerName $target }
            Set-Status "Finalizado" ([System.Drawing.Color]::LightGreen)
        }
        "ChkDsk /r" {
            Invoke-ActionButton -ComputerName $target -UseCancel $false `
                -StatusMsg "Ejecutando ChkDsk /r en '$target'..." `
                -Action    { Invoke-RemoteChkdsk -ComputerName $target }
            Set-Status "Finalizado" ([System.Drawing.Color]::LightGreen)
        }
    }
})

$btnClear.Add_Click({
    $script:outputBox.Clear()
    Set-Status "Listo" $white
})

$btnCancel.Add_Click({
    # Nota: la cancelacion es "entre pasos" — Invoke-Step comprueba $cancelRequested al inicio
    # de cada paso. Un Invoke-Command en curso no se puede interrumpir (WinForms es single-thread).
    # Remove-PSSession cierra sesiones persistentes pero no corta una llamada bloqueante activa.
    $script:cancelRequested = $true
    Write-Warn "Cancelacion solicitada - abortando sesiones remotas..."
    Set-Status "Cancelando..." ([System.Drawing.Color]::Orange)
    Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
    [System.Windows.Forms.Application]::DoEvents()
})

# ── Eventos del panel lateral de equipos ─────────────────────────

$btnAddEquipo.Add_Click({
    $input = Get-Input "Nombre del equipo a anadir:" "Anadir equipo"
    $input = $input.Trim().ToUpper()
    if ([string]::IsNullOrEmpty($input)) { return }
    $exists = $script:flowEquipos.Controls | Where-Object {
        $_ -is [System.Windows.Forms.Panel] -and $_.Tag -eq $input
    }
    if ($exists) {
        [System.Windows.Forms.MessageBox]::Show(
            "'$input' ya esta en la lista.",
            "Duplicado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $card = New-EquipoCard $input
    $script:flowEquipos.Controls.Add($card)
    Save-EquipoList
    Update-EquipoCard $card
})

$btnRemoveEquipo.Add_Click({
    if (-not $script:EquipoSelCard) {
        [System.Windows.Forms.MessageBox]::Show(
            "Haz clic en un equipo de la lista para seleccionarlo primero.",
            "Nada seleccionado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $script:flowEquipos.Controls.Remove($script:EquipoSelCard)
    $script:EquipoSelCard = $null
    Save-EquipoList
})

$btnRefreshEquipos.Add_Click({
    if ($script:flowEquipos.Controls.Count -eq 0) { return }
    Set-Status "Refrescando estado de equipos..." ([System.Drawing.Color]::Yellow)
    Refresh-EquipoEstados
    Set-Status "Listo" $white
})

# Enter en campo equipo -> ping
$txtEquipo.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") { $btnPing.PerformClick() }
})

$form.Add_Shown({
    Append-Output "  Herramienta de Administracion Remota v2.8.3" ([System.Drawing.Color]::FromArgb(0, 190, 255))
    Append-Output "  Accenture / Airbus  |  PowerShell 5.1"    $silver
    Write-Sep
    Append-Output "  > Introduce el nombre del equipo en el campo superior." $silver
    Append-Output "  > Usa 'Ping' para verificar conectividad antes de operar." $silver
    Append-Output "  > Los botones se bloquean mientras hay una tarea en curso." $silver
    Append-Output "  > 'Info del Sistema' muestra SO, IP, MAC, CPU, RAM, discos y mas." $silver
    Write-Sep
    Append-Output "" $white
    # Cargar equipos en seguimiento guardados en sesiones anteriores
    Load-EquipoList
})

[System.Windows.Forms.Application]::Run($form)

#endregion
