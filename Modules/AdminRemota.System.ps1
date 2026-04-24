#region ═══════════════════════════════════════════════════════════
# OPCION 4 - INFORMACION DEL SISTEMA
#═══════════════════════════════════════════════════════════════════

function Invoke-SystemInfo {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Info "Recopilando informacion del sistema '$ComputerName'..."
    Write-Sep

    $info = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
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
# OPCION 6 - INFORMACION DE DRIVERS
#═══════════════════════════════════════════════════════════════════

function Invoke-DriverInfo {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Info "Informacion de drivers de '$ComputerName'..."
    Write-Sep

    $data = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
        $drivers = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
                   Where-Object { $_.DeviceName } |
                   Sort-Object DeviceClass, DeviceName |
                   Select-Object DeviceClass, DeviceName, DriverVersion, DriverDate, Manufacturer

        $problems = @{}
        Get-PnpDevice -ErrorAction SilentlyContinue |
            Where-Object { $_.Status -ne 'OK' -and $_.FriendlyName } |
            ForEach-Object { $problems[$_.FriendlyName] = $_.Status }

        return @{ Drivers = @($drivers); Problems = $problems }
    }

    if (-not $data) {
        Write-Warn "No se pudo obtener informacion de drivers (equipo sin respuesta o WinRM no disponible)."
        return
    }

    $cyan  = [System.Drawing.Color]::FromArgb(0, 190, 255)
    $white = $script:White
    $red   = [System.Drawing.Color]::Tomato

    $catNames = @{
        'Display'         = 'Grafica / Display'
        'Net'             = 'Red (Ethernet / WiFi / LAN)'
        'USB'             = 'USB (Controladores)'
        'HIDClass'        = 'HID (Entrada: raton, teclado, gamepad)'
        'Mouse'           = 'Raton'
        'Keyboard'        = 'Teclado'
        'Media'           = 'Audio / Multimedia'
        'AudioEndpoint'   = 'Audio (endpoints)'
        'HDAUDIO'         = 'Audio HD'
        'HDC'             = 'Controladora de disco (IDE/ATA)'
        'SCSIAdapter'     = 'Controladora SCSI / RAID / NVMe'
        'DiskDrive'       = 'Unidades de disco'
        'Bluetooth'       = 'Bluetooth'
        'System'          = 'Sistema / Placa base'
        'Ports'           = 'Puertos (COM / LPT)'
        'Printer'         = 'Impresoras'
        'Battery'         = 'Bateria'
        'Camera'          = 'Camara / Imaging'
        'Biometric'       = 'Biometrico (huella dactilar)'
        'Monitor'         = 'Monitor'
        'Firmware'        = 'Firmware (UEFI/ACPI)'
        'SmartCardReader' = 'Lector de tarjetas inteligentes'
        'SecurityDevices' = 'Seguridad (TPM)'
        'Processor'       = 'Procesador'
        'Volume'          = 'Volumenes de disco'
        'WPD'             = 'Dispositivos portatiles (MTP)'
        'Sensor'          = 'Sensores'
        'SoftwareDevice'  = 'Dispositivos de software'
        'MTD'             = 'Dispositivos MTD'
        'UCM'             = 'Administrador de conector USB-C'
    }

    # Dispositivos con problemas primero
    if ($data.Problems.Count -gt 0) {
        Append-Output "  [! DISPOSITIVOS CON PROBLEMAS !]" $red
        foreach ($name in ($data.Problems.Keys | Sort-Object)) {
            Append-Output ("    ! {0,-50}  [{1}]" -f $name, $data.Problems[$name]) $red
        }
        Append-Output "" $white
    }

    # Drivers agrupados por categoria
    $grouped = @($data.Drivers) | Group-Object DeviceClass | Sort-Object Name
    foreach ($grp in $grouped) {
        $label = if ($catNames.ContainsKey($grp.Name)) { $catNames[$grp.Name] }
                 elseif ($grp.Name)                     { $grp.Name }
                 else                                   { 'Sin categoria' }
        Append-Output ("  [{0}]" -f $label.ToUpper()) $cyan
        foreach ($d in $grp.Group) {
            $ver  = if ($d.DriverVersion) { $d.DriverVersion } else { '---' }
            $date = try   { ([datetime]$d.DriverDate).ToString('yyyy-MM-dd') }
                    catch { '???' }
            $col  = if ($data.Problems.ContainsKey($d.DeviceName)) { $red } else { $white }
            $pfx  = if ($data.Problems.ContainsKey($d.DeviceName)) { '! ' } else { '  ' }
            Append-Output ("    {0}{1,-50} v{2,-22} {3}" -f $pfx, $d.DeviceName, $ver, $date) $col
        }
        Append-Output "" $white
    }

    Write-Sep
    $total = @($data.Drivers).Count
    $nProb = $data.Problems.Count
    if ($nProb -gt 0) {
        Write-Warn ("Total: {0} drivers  |  {1} dispositivo(s) con problemas detectados" -f $total, $nProb)
    } else {
        Write-Ok ("Total: {0} drivers  |  Sin problemas detectados" -f $total)
    }
    Append-Output "" $white
}

#endregion

#region ═══════════════════════════════════════════════════════════
# OPCION 3 - BORRAR DRIVERS USB
#═══════════════════════════════════════════════════════════════════


function Invoke-UsbDriverClean {
    param([Parameter(Mandatory)][string]$ComputerName)
    # Funcionalidad USB temporalmente deshabilitada. Pendiente de reactivacion.
    Write-Warn "Borrado de drivers USB no disponible en esta version."
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

# Helper: Lee las ultimas $TailLines lineas de un log remoto y devuelve la ultima linea
# que coincida con $SuccessPattern (regex case-insensitive). Solo transfiere escalares por
# WinRM para evitar serializar arrays grandes. Compatible con PS 5.1.
# Devuelve @{ Found=$bool; Line=$string; Source=$string; Details=$string }
function Get-RemoteLogSuccessLine {
    param(
        [Parameter(Mandatory)][string] $ComputerName,
        [Parameter(Mandatory)][string] $LogPath,
        [Parameter(Mandatory)][string] $SuccessPattern,
        [int]                          $TailLines = 100
    )
    try {
        $r = Invoke-LocalOrRemote -ComputerName $ComputerName `
            -ArgumentList $LogPath, $SuccessPattern, $TailLines `
            -ScriptBlock {
                param([string]$path, [string]$pattern, [int]$tail)
                if (-not (Test-Path $path -ErrorAction SilentlyContinue)) {
                    return @{ Found=$false; Line=''; Source=''; Details="Log no encontrado: $path" }
                }
                $lines = Get-Content $path -Tail $tail -ErrorAction SilentlyContinue
                # Select-Object -Last 1: devuelve solo la coincidencia mas reciente
                $hit = @($lines | Where-Object { $_ -match $pattern }) | Select-Object -Last 1
                if ($hit) {
                    return @{ Found=$true; Line=$hit.Trim(); Source=$path; Details='OK' }
                }
                return @{ Found=$false; Line=''; Source=$path; Details="Sin coincidencia (ultimas $tail lineas revisadas)" }
            }
        if ($null -eq $r) { return @{ Found=$false; Line=''; Source=''; Details='Sin respuesta remota' } }
        return $r
    } catch {
        return @{ Found=$false; Line=''; Source=''; Details="Error al leer log: $($_.Exception.Message)" }
    }
}

# ── Reparacion / reinstalacion del cliente SCCM ────────────────────────────────
# Paso 1: ccmrepair.exe  -> reparacion in-place, no destructiva, sincrona (~1-2 min)
# Paso 2: ccmsetup.exe   -> reinstalacion completa, asincrona (lanza el instalador y vuelve)
function Invoke-SccmRepair {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Sep
    Write-Info "SCCM Repair / Reinstall en '$ComputerName'"
    Write-Sep

    # Detectar VPN: WinRM/WMI suele estar bloqueado en segmentos VPN
    $zone = Get-TargetNetworkZone $ComputerName
    if ($zone -eq 'VPN') {
        Write-Warn "Equipo conectado por VPN."
        Write-Info "  La ejecucion remota de ccmrepair/ccmsetup puede no estar disponible (WinRM bloqueado)."
        Write-Info "  Se intentara igualmente si confirmas, pero puede fallar."
        Append-Output "" $script:White
    }

    # ── Paso 1: ccmrepair ────────────────────────────────────────────────────────
    if (Confirm-Action (
        "Paso 1 de 2 - Reparacion del cliente SCCM en '$ComputerName':`n`n" +
        "  Ejecutara:  C:\Windows\CCM\ccmrepair.exe`n`n" +
        "  - Repara el cliente SCCM sin desinstalarlo.`n" +
        "  - No destructivo: mantiene politicas y cache.`n" +
        "  - Duracion estimada: 1-3 minutos.`n`n" +
        "¿Ejecutar ccmrepair ahora?"
    ) "SCCM Repair") {

        Write-Info "[1/2] Verificando ruta C:\Windows\CCM\ en '$ComputerName'..."
        Set-Status "Verificando CCM en '$ComputerName'..." ([System.Drawing.Color]::Yellow)
        [System.Windows.Forms.Application]::DoEvents()

        $pathCheck = $null
        try {
            $pathCheck = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
                return Test-Path 'C:\Windows\CCM\ccmrepair.exe'
            }
        } catch {
            Write-Fail "No se pudo verificar la ruta en '$ComputerName': $($_.Exception.Message)"
            if ($zone -eq 'VPN') {
                Write-Info "  Causa probable: WinRM bloqueado por VPN/firewall."
            }
            $pathCheck = $false
        }

        if (-not $pathCheck) {
            Write-Fail "C:\Windows\CCM\ccmrepair.exe no encontrado en '$ComputerName'."
            Write-Info "  El cliente SCCM puede no estar instalado o la ruta es diferente."
        } else {
            Write-Ok   "  ccmrepair.exe localizado."
            Write-Info "[1/2] Ejecutando ccmrepair.exe en '$ComputerName'..."
            Write-Info "  Espera - puede tardar 1-3 minutos..."
            Set-Status "ccmrepair en curso en '$ComputerName'..." ([System.Drawing.Color]::Yellow)
            [System.Windows.Forms.Application]::DoEvents()

            $repairResult = $null
            try {
                # ccmrepair es sincrono: esperar hasta 5 minutos antes de timeout
                $sessOpt = New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 300000
                $repairResult = Invoke-Command -ComputerName $ComputerName `
                    -SessionOption $sessOpt -ErrorAction Stop -ScriptBlock {
                        try {
                            $proc = Start-Process -FilePath 'C:\Windows\CCM\ccmrepair.exe' `
                                                  -Wait -PassThru -ErrorAction Stop
                            return @{ ExitCode = $proc.ExitCode; Error = $null }
                        } catch {
                            return @{ ExitCode = -1; Error = $_.Exception.Message }
                        }
                    }
            } catch {
                Write-Fail "Error al ejecutar ccmrepair remotamente: $($_.Exception.Message)"
                if ($zone -eq 'VPN') {
                    Write-Info "  Causa probable: WinRM bloqueado por VPN/firewall."
                }
            }

            if ($repairResult) {
                if ($repairResult.Error) {
                    Write-Fail "ccmrepair fallo: $($repairResult.Error)"
                    Set-Status "ccmrepair ERROR en '$ComputerName'" ([System.Drawing.Color]::Tomato)
                } elseif ($repairResult.ExitCode -eq 0) {
                    Write-Ok  "[1/2] ccmrepair completado correctamente. ExitCode=0"
                    Set-Status "ccmrepair OK en '$ComputerName'" ([System.Drawing.Color]::LightGreen)
                    # Validar resultado por log remoto (tail corto; no se vuelca el log completo)
                    # Log tipico: C:\Windows\CCM\Logs\ccmrepair.log (formato CMTrace)
                    $logR = Get-RemoteLogSuccessLine -ComputerName $ComputerName `
                                -LogPath       'C:\Windows\CCM\Logs\ccmrepair.log' `
                                -SuccessPattern 'repair.*succeed|succeeded|CcmRepair.*complet|Repair.*complet' `
                                -TailLines     100
                    if     ($logR.Found)  { Write-Ok   "  Log ccmrepair: $($logR.Line)" }
                    elseif ($logR.Source) { Write-Warn "  Log ccmrepair: sin linea concluyente ($($logR.Details))." }
                    else                  { Write-Warn "  Log ccmrepair: $($logR.Details)" }
                } else {
                    Write-Warn "[1/2] ccmrepair finalizo con ExitCode=$($repairResult.ExitCode)."
                    Write-Info "  ExitCode distinto de 0 puede indicar reinicio pendiente o aviso menor."
                    Set-Status "ccmrepair WARN en '$ComputerName'" ([System.Drawing.Color]::Yellow)
                }
            }
        }
        Append-Output "" $script:White
    }

    # ── Paso 2: ccmsetup (reinstalacion completa) ─────────────────────────────────
    Write-Sep
    if (Confirm-Action (
        "Paso 2 de 2 - Reinstalacion completa del cliente SCCM en '$ComputerName':`n`n" +
        "  Ejecutara:  C:\Windows\CCMSetup\ccmsetup.exe`n`n" +
        "  - Reinstala el cliente SCCM desde cero.`n" +
        "  - MAS agresivo que ccmrepair: desinstala y vuelve a instalar.`n" +
        "  - El proceso se lanza en segundo plano en el equipo remoto.`n" +
        "  - Puede tardar 5-15 minutos en completarse.`n" +
        "  - Monitorizar progreso en: C:\Windows\CCMSetup\Logs\ccmsetup.log`n`n" +
        "¿Ejecutar ccmsetup ahora?"
    ) "SCCM Reinstall") {

        Write-Info "[2/2] Verificando ruta C:\Windows\CCMSetup\ en '$ComputerName'..."
        Set-Status "Verificando CCMSetup en '$ComputerName'..." ([System.Drawing.Color]::Yellow)
        [System.Windows.Forms.Application]::DoEvents()

        $pathCheck2 = $null
        try {
            $pathCheck2 = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
                return Test-Path 'C:\Windows\CCMSetup\ccmsetup.exe'
            }
        } catch {
            Write-Fail "No se pudo verificar la ruta en '$ComputerName': $($_.Exception.Message)"
            if ($zone -eq 'VPN') {
                Write-Info "  Causa probable: WinRM bloqueado por VPN/firewall."
            }
            $pathCheck2 = $false
        }

        if (-not $pathCheck2) {
            Write-Fail "C:\Windows\CCMSetup\ccmsetup.exe no encontrado en '$ComputerName'."
            Write-Info "  El instalador SCCM puede no estar cacheado en el equipo."
            Write-Info "  Alternativa: forzar descarga via Software Center o consola SCCM."
        } else {
            Write-Ok   "  ccmsetup.exe localizado."
            Write-Info "[2/2] Lanzando ccmsetup.exe en '$ComputerName' (proceso en segundo plano)..."
            Set-Status "Lanzando ccmsetup en '$ComputerName'..." ([System.Drawing.Color]::Yellow)
            [System.Windows.Forms.Application]::DoEvents()

            $setupResult = $null
            try {
                $setupResult = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
                    try {
                        # ccmsetup es asincrono: Start-Process sin -Wait intencionado
                        Start-Process -FilePath 'C:\Windows\CCMSetup\ccmsetup.exe' `
                                      -ErrorAction Stop | Out-Null
                        return @{ OK = $true; Error = $null }
                    } catch {
                        return @{ OK = $false; Error = $_.Exception.Message }
                    }
                }
            } catch {
                Write-Fail "Error al lanzar ccmsetup remotamente: $($_.Exception.Message)"
                if ($zone -eq 'VPN') {
                    Write-Info "  Causa probable: WinRM bloqueado por VPN/firewall."
                }
            }

            if ($setupResult) {
                if ($setupResult.Error) {
                    Write-Fail "[2/2] ccmsetup no pudo iniciarse: $($setupResult.Error)"
                    Set-Status "ccmsetup ERROR en '$ComputerName'" ([System.Drawing.Color]::Tomato)
                } else {
                    Write-Ok  "[2/2] ccmsetup lanzado correctamente en '$ComputerName'."
                    Write-Info "  La instalacion continua en segundo plano en el equipo remoto."
                    Set-Status "ccmsetup lanzado en '$ComputerName'" ([System.Drawing.Color]::LightGreen)
                    # Comprobacion inmediata del log (ccmsetup es asincrono; puede no haber linea aun).
                    # Log tipico: C:\Windows\CCMSetup\Logs\ccmsetup.log (formato CMTrace)
                    $logS = Get-RemoteLogSuccessLine -ComputerName $ComputerName `
                                -LogPath        'C:\Windows\CCMSetup\Logs\ccmsetup.log' `
                                -SuccessPattern 'CcmSetup.*exiting.*return code 0|Installation succeeded|CcmSetup succeeded' `
                                -TailLines      100
                    if     ($logS.Found)  { Write-Ok   "  Log ccmsetup: $($logS.Line)" }
                    elseif ($logS.Source) { Write-Info "  Log ccmsetup: instalacion en curso, sin confirmacion aun. Log: $($logS.Source)" }
                    else                  { Write-Warn "  Log ccmsetup: $($logS.Details)" }
                }
            }
        }
        Append-Output "" $script:White
    }

    Write-Sep
    Append-Output "" $script:White
}

#endregion
