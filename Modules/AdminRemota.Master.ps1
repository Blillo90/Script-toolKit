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
            $result = Invoke-LocalOrRemote -ComputerName $script:Target -ArgumentList (,$certDefsForRemote) -ScriptBlock {
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
                            # $script:CertreqEnrollBlock se pasa como -ScriptBlock (no como ArgumentList)
                            # para evitar que PSRP serialice el ScriptBlock a string al cruzar WinRM.
                            # OperationTimeoutMs=180000 (3 min): PASO 2 puede tardar hasta
                            # cesTimeout(20s) * nUrls(2) = 40s en el peor caso.
                            $remoteResult = Invoke-LocalOrRemote -ComputerName $script:Target `
                                -ArgumentList $urls, $ct `
                                -OperationTimeoutMs 180000 `
                                -ScriptBlock $script:CertreqEnrollBlock
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
            Invoke-LocalOrRemote -ComputerName $script:Target -ArgumentList $script:ExpectedIssuerLike -ScriptBlock {
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
        Invoke-LocalOrRemote -ComputerName $script:Target -ScriptBlock {
            $f = Get-Item "C:\success.txt" -ErrorAction SilentlyContinue
            if (-not $f) { return @{ Status="ERROR"; Details="No existe C:\success.txt" } }
            # Semana actual: lunes a domingo (semana ISO). DayOfWeek: 0=Dom, 1=Lun ... 6=Sab.
            $hoy = Get-Date
            $dia = [int]$hoy.DayOfWeek
            if ($dia -eq 0) { $dia = 7 }           # Tratar domingo como dia 7 (lunes = dia 1)
            $inicioSemana = $hoy.Date.AddDays(1 - $dia)
            $finSemana    = $inicioSemana.AddDays(7)
            $s = if ($f.LastWriteTime -ge $inicioSemana -and $f.LastWriteTime -lt $finSemana) { "OK" } else { "WARN" }
            $rango = "$($inicioSemana.ToString('dd/MM')) - $($finSemana.AddDays(-1).ToString('dd/MM/yyyy'))"
            return @{ Status=$s; Details="Fecha: $($f.LastWriteTime) | Semana: $rango" }
        }
    }

    # Step 4: Ciclos SCCM
    Invoke-Step -Name "SCCM Client Cycles" -ScriptBlock {
        $res = Invoke-LocalOrRemote -ComputerName $script:Target -ScriptBlock $script:SccmCyclesBlock
        if ($res -and $res.Steps) { Write-StepList $res.Steps }
        if ($res) { return @{ Status=$res.Status; Details="" } }
        return $res
    }

    # Step 5: Centro de Software
    # Solo se considera positivo si existen aplicaciones con 'Install' en AllowedActions.
    # Eso excluye Windows Updates, Feature Updates desplegados como app, y apps ya instaladas.
    Invoke-Step -Name "Centro de Software (CCM_Application)" -ScriptBlock {
        # Detectar VPN antes de intentar WMI/WinRM remoto: en VPN esos puertos suelen estar
        # bloqueados por firewall de segmentacion -> evitar error generico, mostrar aviso claro.
        $zone = Get-TargetNetworkZone $script:Target
        if ($zone -eq 'VPN') {
            return @{
                Status  = "WARN"
                Details = "Equipo conectado por VPN. La comprobacion SCCM remota no esta disponible en este segmento de red (WMI/WinRM bloqueado por firewall)."
            }
        }

        $diag = Test-RemoteSccmReady -ComputerName $script:Target
        if ($diag -ne "OK") { return @{ Status="ERROR"; Details=$diag } }

        Invoke-LocalOrRemote -ComputerName $script:Target -ScriptBlock {
            $allCcm = @(Get-CimInstance -Namespace "root\ccm\ClientSDK" -ClassName "CCM_Application" `
                                        -ErrorAction SilentlyContinue)

            # Aplicaciones realmente disponibles para instalar:
            # AllowedActions debe contener 'Install' -> no instaladas y el usuario puede instalarlas.
            # Updates de Windows / Feature Updates no tienen 'Install' en AllowedActions.
            $installable = @($allCcm | Where-Object { $_.AllowedActions -contains 'Install' })

            if ($installable.Count -gt 0) {
                $ver     = (Get-CimInstance -Namespace "root\ccm" -ClassName "SMS_Client" `
                                            -ErrorAction SilentlyContinue).ClientVersion
                $ejemplos = ($installable | Select-Object -First 3 |
                             ForEach-Object { $_.Name }) -join ' | '
                return @{
                    Status  = "OK"
                    Details = ("Aplicaciones disponibles para instalar: {0} | ClientVersion={1} | Ejemplos: {2}" `
                               -f $installable.Count, $ver, $ejemplos)
                }
            }

            # Habia contenido pero ninguno con accion Install (solo updates u otros)
            if ($allCcm.Count -gt 0) {
                return @{
                    Status  = "WARN"
                    Details = ("Sin aplicaciones disponibles para instalar. " +
                               "Solo hay actualizaciones u otros elementos ({0} entradas en CCM_Application)." `
                               -f $allCcm.Count)
                }
            }

            return @{ Status="WARN"; Details="CcmExec Running y namespace OK, pero sin contenido en CCM_Application" }
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
            $found = Invoke-LocalOrRemote -ComputerName $script:Target -ArgumentList (,$checkFilters) -ScriptBlock {
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
