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

    # Detectar VPN: avisar antes de intentar WMI/WinRM que pueden estar bloqueados.
    $zone = Get-TargetNetworkZone $ComputerName
    if ($zone -eq 'VPN') {
        Write-Warn "Equipo conectado por VPN. La comprobacion SCCM remota puede no estar disponible."
        Write-Info "  WMI y WinRM suelen estar bloqueados en segmentos VPN corporativos."
        Write-Info "  Se intentara igualmente, pero puede fallar."
    }

    # Diagnostico previo (reutiliza helper comun)
    $diag = Test-RemoteSccmReady -ComputerName $ComputerName
    if ($diag -ne "OK") {
        if ($zone -eq 'VPN') {
            Write-Fail "No se puede acceder al cliente SCCM remotamente (posible VPN o firewall)."
            Write-Info "  Detalle tecnico: $diag"
        } else {
            Write-Fail "No se puede consultar el Centro de Software: $diag"
        }
        return
    }

    $apps = $null
    try {
        # Usar Invoke-LocalOrRemote (WinRM) en lugar de Get-CimInstance -ComputerName (DCOM/CIM)
        # para mantener coherencia con Invoke-MasterCheck y el helper Test-RemoteSccmReady.
        $apps = Invoke-LocalOrRemote -ComputerName $ComputerName -ArgumentList $appName -ScriptBlock {
            param($name)
            Get-CimInstance -Namespace "root\ccm\ClientSDK" -ClassName "CCM_Application" `
                            -ErrorAction Stop |
            Where-Object { $_.Name -like "*$name*" }
        }
    } catch {
        if ($zone -eq 'VPN') {
            Write-Fail "No se puede acceder al cliente SCCM remotamente (posible VPN o firewall)."
            Write-Info "  Detalle tecnico: $($_.Exception.Message)"
        } else {
            Write-Fail "Error al consultar CCM_Application: $($_.Exception.Message)"
        }
        return
    }

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
                        $rv = Invoke-LocalOrRemote -ComputerName $ComputerName `
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
    $localMatches = Invoke-LocalOrRemote -ComputerName $ComputerName -ArgumentList $appName -ScriptBlock {
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
        $res = Invoke-LocalOrRemote -ComputerName $script:Target -ScriptBlock $script:SccmCyclesBlock
        if ($res -and $res.Steps) { Write-StepList $res.Steps }
        if ($res) { return @{ Status=$res.Status; Details="" } }
        return $res
    }

    if (Confirm-Action "Ejecutar gpupdate /force en '$ComputerName'?") {
        Invoke-Step -Name "GPUPDATE /force" -ScriptBlock {
            Invoke-RemoteGpupdate -ComputerName $script:Target
        }
    }

    Show-Summary
}

#endregion
