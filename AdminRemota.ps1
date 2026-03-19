#Requires -Version 5.1
<#
.SYNOPSIS
    Herramienta de administracion remota unificada v2.1 (GUI)
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
    2.1
#>

[CmdletBinding()]
param(
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
$script:cancelRequested = $false
$script:Modo            = "Nacional"   # "Nacional" | "Divisional"
$script:Target          = ""
$script:StepResults     = New-Object System.Collections.Generic.List[object]

# Colores GUI (definidos aqui para que esten disponibles en toda la sesion)
$script:White  = [System.Drawing.Color]::White
$script:Silver = [System.Drawing.Color]::Silver

# URLs CES Kerberos centralizadas (unica fuente de verdad)
$script:CesMap = @{
    "Breguet G1"  = @(
        "https://aefews01.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20Breguet%20G1_CES_Kerberos/service.svc/CES",
        "https://aefews02.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20Breguet%20G1_CES_Kerberos/service.svc/CES"
    )
    "da Vinci G1" = @(
        "https://aefews01.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20da%20Vinci%20G1_CES_Kerberos/service.svc/CES",
        "https://aefews02.autoenroll.pki.intra.corp/Airbus%20Issuing%20CA%20da%20Vinci%20G1_CES_Kerberos/service.svc/CES"
    )
}

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
        $status  = if ($res -is [hashtable] -and $res.Status)  { $res.Status }                        else { "OK"  }
        $details = if ($res -is [hashtable] -and $res.Details) { ($res.Details | Out-String).Trim() } else { ""   }
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

# ── Busca software por nombre en el registro Uninstall (32 y 64 bits) ─────────
function Get-RemoteInstalledSoftware {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][string]$NameFilter
    )
    return Invoke-Command -ComputerName $ComputerName -ArgumentList $NameFilter -ScriptBlock {
        param($filter)
        $paths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        foreach ($p in $paths) {
            $match = Get-ItemProperty $p -ErrorAction SilentlyContinue |
                     Where-Object { $_.DisplayName -like $filter } |
                     Select-Object -First 1
            if ($match) { return "$($match.DisplayName) v$($match.DisplayVersion)" }
        }
        return $null
    }
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

            $result = Invoke-Command -ComputerName $script:Target -ScriptBlock {
                $all     = Get-ChildItem "Cert:\LocalMachine\My" -ErrorAction SilentlyContinue
                $now     = Get-Date
                $details = @(); $missing = @(); $cleaned = @()

                foreach ($caEntry in @(
                    @{ Name="Breguet G1";  Filter="*Breguet*" },
                    @{ Name="da Vinci G1"; Filter="*Vinci*"   }
                )) {
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
                    $cesMapSnapshot = $script:CesMap   # captura para closure remoto
                    foreach ($certType in $missingCerts) {
                        $urls = $cesMapSnapshot[$certType]
                        $ct   = $certType
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

    # ── Fase A: Deteccion ────────────────────────────────────────────
    Write-Info "Buscando drivers USB en '$ComputerName'..."
    Write-Sep

    $drivers    = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        Get-PnpDevice | Where-Object {
            $_.FriendlyName -match "USB" -and $_.Class -match "USB"
        } | Sort-Object Class, FriendlyName |
            Select-Object Status, Class, FriendlyName, InstanceId
    }
    $driverList = @($drivers)
    $total      = $driverList.Count

    if ($total -eq 0) {
        Write-Warn "No se encontraron drivers USB en '$ComputerName'."
        Write-Sep
        return
    }

    Write-Info "Se encontraron $total driver(s) USB candidato(s):"
    Append-Output "" $script:White

    $idx = 0
    foreach ($d in $driverList) {
        $idx++
        $stateColor = switch ($d.Status) {
            "OK"      { [System.Drawing.Color]::LightGreen              }
            "Error"   { [System.Drawing.Color]::Tomato                  }
            "Unknown" { [System.Drawing.Color]::Gray                    }
            default   { [System.Drawing.Color]::LightYellow             }
        }
        Append-Output ("    [{0:D2}] {1,-10}  {2,-26}  {3}" -f $idx, $d.Status, $d.Class, $d.FriendlyName) $stateColor
        Append-Output ("           InstanceId: {0}" -f $d.InstanceId) ([System.Drawing.Color]::FromArgb(120, 120, 120))
    }
    Append-Output "" $script:White

    # ── Fase B: Confirmacion ─────────────────────────────────────────
    if (-not (Confirm-Action (
        "Se van a eliminar $total driver(s) USB en '$ComputerName'.`n`n" +
        "Esta operacion puede requerir reinicio y no es facilmente reversible.`n`n" +
        "¿Confirmar borrado?"
    ))) {
        Write-Warn "Borrado cancelado por el usuario. No se ha eliminado nada."
        return
    }

    # ── Fase C: Borrado con progreso visible ─────────────────────────
    Write-Sep
    Write-Info "Iniciando borrado de $total driver(s)..."
    Append-Output "" $script:White

    $okCount    = 0
    $warnCount  = 0
    $errorCount = 0
    $results    = [System.Collections.Generic.List[object]]::new()
    $idx        = 0

    foreach ($d in $driverList) {
        $idx++
        Write-Info ("  [{0}/{1}] Procesando: {2}" -f $idx, $total, $d.FriendlyName)

        $rem = Invoke-Command -ComputerName $ComputerName -ArgumentList $d.InstanceId -ScriptBlock {
            param($instanceId)
            $out = pnputil /remove-device "$instanceId" 2>&1
            return @{ Output = ($out -join ' ').Trim(); ExitCode = $LASTEXITCODE }
        }

        switch ($rem.ExitCode) {
            0 {
                Write-Ok   ("  [{0}/{1}] OK     -> {2}" -f $idx, $total, $d.FriendlyName)
                $okCount++
                $results.Add([PSCustomObject]@{ Status="OK";    Name=$d.FriendlyName; Detail="OK" })
            }
            3010 {
                # pnputil: exito pero se requiere reinicio para completar
                Write-Warn ("  [{0}/{1}] WARN   -> {2}  (requiere reinicio)" -f $idx, $total, $d.FriendlyName)
                $warnCount++
                $results.Add([PSCustomObject]@{ Status="WARN";  Name=$d.FriendlyName; Detail="Eliminado, pendiente reinicio" })
            }
            default {
                Write-Fail ("  [{0}/{1}] ERROR  -> {2}  | ExitCode={3} | {4}" -f $idx, $total, $d.FriendlyName, $rem.ExitCode, $rem.Output)
                $errorCount++
                $results.Add([PSCustomObject]@{ Status="ERROR"; Name=$d.FriendlyName; Detail="ExitCode=$($rem.ExitCode)" })
            }
        }
    }

    # ── Fase D: Resumen final ────────────────────────────────────────
    Append-Output "" $script:White
    Write-Sep
    Write-Info "RESUMEN DE BORRADO USB"
    Write-Sep
    Append-Output ("  Total encontrados  : {0}" -f $total)      $script:White
    Append-Output ("  Eliminados OK      : {0}" -f $okCount)    ([System.Drawing.Color]::LightGreen)
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

    # Ofrecer reinicio si hubo al menos un borrado (OK o pendiente de reinicio)
    $deleted = $okCount + $warnCount
    if ($deleted -gt 0) {
        $reinicioMsg = "Se procesaron correctamente $deleted de $total driver(s)."
        if ($warnCount -gt 0) { $reinicioMsg += "`n$warnCount requieren reinicio para completarse." }
        $reinicioMsg += "`n`n¿Reiniciar '$ComputerName' ahora?"
        if (Confirm-Action $reinicioMsg) {
            Restart-Computer -ComputerName $ComputerName -Force
            Write-Ok "Reiniciando '$ComputerName'..."
        }
    }
}

#endregion

#region ═══════════════════════════════════════════════════════════
# INTERFAZ GRAFICA (WinForms)
#═══════════════════════════════════════════════════════════════════

$script:ExpectedIssuerLike = $ExpectedIssuerLike

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
$form.Size            = New-Object System.Drawing.Size(920, 680)
$form.MinimumSize     = New-Object System.Drawing.Size(720, 500)
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
$topPanel.Height    = 192
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

# ── Fila 2: acciones rapidas ──────────────────────────────────────
$btnRestart      = New-FlatButton "  Reiniciar"         10 153 140 28 ([System.Drawing.Color]::FromArgb(160,  80,  0))
$btnGPUpdate     = New-FlatButton "  GPUpdate /force"  155 153 155 28 ([System.Drawing.Color]::FromArgb(  0, 100, 140))
$btnPolicyCycles = New-FlatButton "  Ciclos SCCM"      315 153 150 28 ([System.Drawing.Color]::FromArgb( 60,  80, 130))

$btnCancel.Enabled = $false

foreach ($b in @($btnMaster,$btnSoftware,$btnInfo,$btnUsb,$btnClear,$btnCancel,$btnRestart,$btnGPUpdate,$btnPolicyCycles)) {
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
$script:outputBox.BringToFront()   # Fill solo ocupa el espacio restante

$script:statusLabel           = New-Object System.Windows.Forms.Label
$script:statusLabel.Text      = "  Listo - introduce un equipo y pulsa una accion"
$script:statusLabel.Dock      = "Fill"
$script:statusLabel.ForeColor = $white
$script:statusLabel.Font      = $fontSmall
$script:statusLabel.TextAlign = "MiddleLeft"
$statusBar.Controls.Add($script:statusLabel)

# ── Helpers de control de UI ──────────────────────────────────────

# Lista de botones de accion (todos excepto Cancelar y Limpiar)
$script:ActionButtons = @($btnMaster,$btnSoftware,$btnInfo,$btnUsb,$btnRestart,$btnGPUpdate,$btnPolicyCycles,$btnPing)

function Set-ButtonsEnabled {
    param([bool]$Enabled)
    foreach ($b in $script:ActionButtons) { $b.Enabled = $Enabled }
    $btnCancel.Enabled = -not $Enabled
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
    Set-Status "Haciendo ping a '$computer'..." ([System.Drawing.Color]::Yellow)
    Write-Sep
    if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
        Set-Status "  '$computer' esta ONLINE" ([System.Drawing.Color]::LightGreen)
        Write-Ok "Ping OK -> '$computer' accesible."
    } else {
        Set-Status "  '$computer' NO accesible" ([System.Drawing.Color]::Tomato)
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

$btnGPUpdate.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    Set-ButtonsEnabled $false
    Set-Status "Ejecutando gpupdate /force en '$target'..." ([System.Drawing.Color]::Yellow)
    Write-Sep
    Write-Info "gpupdate /force en '$target'..."
    try {
        $res = Invoke-RemoteGpupdate -ComputerName $target
        if ($res.Status -eq "OK") {
            Write-Ok "gpupdate completado correctamente ($($res.Details))."
            Set-Status "GPUpdate OK en '$target'" ([System.Drawing.Color]::LightGreen)
        } else {
            Write-Warn "gpupdate: $($res.Details)."
            Set-Status "GPUpdate WARN ($($res.Details))" ([System.Drawing.Color]::Yellow)
        }
    } catch {
        Write-Fail "Error ejecutando gpupdate en '$target': $_"
        Set-Status "Error en gpupdate" ([System.Drawing.Color]::Tomato)
    }
    Write-Sep
    Append-Output "" $white
    Set-ButtonsEnabled $true
})

$btnPolicyCycles.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    Set-ButtonsEnabled $false
    Set-Status "Lanzando ciclos SCCM en '$target'..." ([System.Drawing.Color]::Yellow)
    Write-Sep
    Write-Info "Ciclos de politicas SCCM en '$target'..."
    try {
        $result = Invoke-Command -ComputerName $target -ScriptBlock $script:SccmCyclesBlock
        switch ($result.Status) {
            "OK"    { Write-Ok   "Ciclos completados: $($result.Details)"
                      Set-Status "Ciclos SCCM OK en '$target'" ([System.Drawing.Color]::LightGreen) }
            "WARN"  { Write-Warn "Ciclos con avisos: $($result.Details)"
                      Set-Status "Ciclos SCCM con avisos" ([System.Drawing.Color]::Yellow) }
            "ERROR" { Write-Fail "Ciclos con errores: $($result.Details)"
                      Set-Status "Error en ciclos SCCM" ([System.Drawing.Color]::Tomato) }
        }
    } catch {
        Write-Fail "Error al lanzar ciclos SCCM en '$target': $_"
        Write-Info "  Verifica que el cliente SCCM esta instalado y activo."
        Set-Status "Error en ciclos SCCM" ([System.Drawing.Color]::Tomato)
    }
    Write-Sep
    Append-Output "" $white
    Set-ButtonsEnabled $true
})

$btnClear.Add_Click({
    $script:outputBox.Clear()
    Set-Status "Listo" $white
})

$btnCancel.Add_Click({
    $script:cancelRequested = $true
    Write-Warn "Cancelacion solicitada - abortando sesiones remotas..."
    Set-Status "Cancelando..." ([System.Drawing.Color]::Orange)
    Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
    [System.Windows.Forms.Application]::DoEvents()
})

# Enter en campo equipo -> ping
$txtEquipo.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") { $btnPing.PerformClick() }
})

$form.Add_Shown({
    Append-Output "  Herramienta de Administracion Remota v2.1" ([System.Drawing.Color]::FromArgb(0, 190, 255))
    Append-Output "  Accenture / Airbus  |  PowerShell 5.1"    $silver
    Write-Sep
    Append-Output "  > Introduce el nombre del equipo en el campo superior." $silver
    Append-Output "  > Usa 'Ping' para verificar conectividad antes de operar." $silver
    Append-Output "  > Los botones se bloquean mientras hay una tarea en curso." $silver
    Append-Output "  > 'Info del Sistema' muestra SO, IP, MAC, CPU, RAM, discos y mas." $silver
    Write-Sep
    Append-Output "" $white
})

[System.Windows.Forms.Application]::Run($form)

#endregion
