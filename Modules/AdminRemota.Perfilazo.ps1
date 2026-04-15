#region ═══════════════════════════════════════════════════════════
# PERFILAZO - BACKUP Y BORRADO DE PERFIL DE USUARIO
#═══════════════════════════════════════════════════════════════════

# Muestra un selector modal con los perfiles obtenidos de Win32_UserProfile.
# Devuelve el LocalPath real del perfil elegido, o $null si se cancela.
function Show-ProfilePicker {
    param(
        [Parameter(Mandatory)][object[]]$Profiles,
        [string]$Title = "Seleccionar perfil"
    )

    $pf = New-Object System.Windows.Forms.Form
    $pf.Text            = $Title
    $pf.Size            = New-Object System.Drawing.Size(520, 148)
    $pf.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $pf.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $pf.MaximizeBox     = $false
    $pf.MinimizeBox     = $false

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = "Selecciona el perfil a tratar:"
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size     = New-Object System.Drawing.Size(480, 18)
    $pf.Controls.Add($lbl)

    $cbo              = New-Object System.Windows.Forms.ComboBox
    $cbo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cbo.Location     = New-Object System.Drawing.Point(12, 34)
    $cbo.Size         = New-Object System.Drawing.Size(480, 24)
    foreach ($p in $Profiles) {
        $loadStr = if ($p.Loaded) { 'Loaded=True ' } else { 'Loaded=False' }
        $null = $cbo.Items.Add(("{0,-22}  |  {1}  |  {2}" -f $p.Name, $loadStr, $p.LocalPath))
    }
    $cbo.SelectedIndex = 0
    $pf.Controls.Add($cbo)

    $btnOk            = New-Object System.Windows.Forms.Button
    $btnOk.Text       = "Aceptar"
    $btnOk.Location   = New-Object System.Drawing.Point(304, 72)
    $btnOk.Size       = New-Object System.Drawing.Size(88, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $pf.AcceptButton  = $btnOk
    $pf.Controls.Add($btnOk)

    $btnCnl           = New-Object System.Windows.Forms.Button
    $btnCnl.Text      = "Cancelar"
    $btnCnl.Location  = New-Object System.Drawing.Point(400, 72)
    $btnCnl.Size      = New-Object System.Drawing.Size(92, 28)
    $btnCnl.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $pf.CancelButton  = $btnCnl
    $pf.Controls.Add($btnCnl)

    $dlgResult = $pf.ShowDialog()
    $selIdx    = $cbo.SelectedIndex
    $pf.Dispose()

    if ($dlgResult -ne [System.Windows.Forms.DialogResult]::OK -or $selIdx -lt 0) { return $null }
    return $Profiles[$selIdx].LocalPath
}

function Invoke-Perfilazo {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Sep
    Write-Info "Perfilazo en '$ComputerName'"
    Write-Sep
    Append-Output "" $script:White

    # ── Listar perfiles disponibles ──────────────────────────────────
    Write-Info "Perfiles locales en '$ComputerName':"
    $perfiles = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
        $excluir = @('Public', 'Default', 'Default User', 'All Users')
        Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue |
            Where-Object {
                (-not $_.Special) -and
                ($_.LocalPath -like 'C:\Users\*') -and
                ((Split-Path $_.LocalPath -Leaf) -notin $excluir)
            } |
            Sort-Object LastUseTime -Descending |
            ForEach-Object {
                $lastUse = if ($_.LastUseTime) { $_.LastUseTime.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
                @{ Name=(Split-Path $_.LocalPath -Leaf); LocalPath=$_.LocalPath; Loaded=$_.Loaded; LastUse=$lastUse }
            }
    }

    if ($perfiles) {
        $idx = 0
        foreach ($p in @($perfiles)) {
            $idx++
            $loadColor = if ($p.Loaded) { [System.Drawing.Color]::Yellow } else { [System.Drawing.Color]::LightGreen }
            $loadStr   = if ($p.Loaded) { 'Loaded=True ' } else { 'Loaded=False' }
            Append-Output ("  [{0:D2}] {1,-20} | {2} | {3} | LastUse={4}" -f `
                $idx, $p.Name, $p.LocalPath, $loadStr, $p.LastUse) $loadColor
        }
    } else {
        Write-Warn "No se pudieron obtener perfiles de '$ComputerName' (o no hay perfiles locales)."
        Write-Sep; Append-Output "" $script:White; return
    }
    Append-Output "" $script:White

    # ── Seleccionar perfil ───────────────────────────────────────────
    $profilePath = Show-ProfilePicker -Profiles @($perfiles) `
                       -Title "Perfilazo - '$ComputerName'"
    if (-not $profilePath) { Write-Warn "Cancelado."; return }
    $usuario = Split-Path $profilePath -Leaf   # solo para mensajes

    $destBase = (Get-Input "Ruta base del share de backups:" "Perfilazo - Destino" "C:\Share").Trim()
    if ([string]::IsNullOrWhiteSpace($destBase)) { Write-Warn "Cancelado."; return }

    # ── Elegir modo: backup nuevo (seguro) o sin backup (peligroso) ──
    $modeChoice = [System.Windows.Forms.MessageBox]::Show(
        "Crear nuevo backup antes de borrar el perfil?`n`n" +
        "SI  = Crear backup completo (recomendado)`n" +
        "NO  = Borrar SIN nuevo backup (PELIGROSO)",
        "Perfilazo - Modo de operacion",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    $skipBackup = ($modeChoice -eq [System.Windows.Forms.DialogResult]::No)

    $rutaExtra = if (-not $skipBackup) {
        (Get-Input "Ruta extra opcional (vaciar si no aplica):" "Perfilazo - Extra" "").Trim()
    } else { "" }

    # ── MODO PELIGROSO: verificar backup previo y confirmar borrado ───
    if ($skipBackup) {
        $prevBackup = $null
        if (Test-Path $destBase) {
            # Intento 1: coincidencia exacta equipo + usuario
            $prevBackup = Get-ChildItem -Path $destBase -Directory -ErrorAction SilentlyContinue |
                          Where-Object { $_.Name -like "${ComputerName}_${usuario}_*" } |
                          Sort-Object LastWriteTime -Descending | Select-Object -First 1
            # Intento 2: solo por equipo si no hay coincidencia exacta
            if (-not $prevBackup) {
                $prevBackup = Get-ChildItem -Path $destBase -Directory -ErrorAction SilentlyContinue |
                              Where-Object { $_.Name -like "${ComputerName}_*" } |
                              Sort-Object LastWriteTime -Descending | Select-Object -First 1
            }
        }

        Write-Warn "[SIN BACKUP] Modo peligroso activado."
        Write-Info "Buscando backup previo en: $destBase"
        if ($prevBackup) {
            Write-Ok "[SIN BACKUP] Backup previo encontrado:"
            Append-Output ("  Ruta:   " + $prevBackup.FullName)                                   ([System.Drawing.Color]::LightGreen)
            Append-Output ("  Nombre: " + $prevBackup.Name)                                       ([System.Drawing.Color]::LightGreen)
            Append-Output ("  Fecha:  " + $prevBackup.LastWriteTime.ToString('yyyy-MM-dd HH:mm')) ([System.Drawing.Color]::LightGreen)
            Append-Output "" $script:White
            Write-Warn "[SIN BACKUP] Este backup puede ser antiguo y NO refleja el estado actual del perfil."
            Write-Warn "[SIN BACKUP] Borrar sin backup nuevo sigue siendo arriesgado."
        } else {
            Write-Fail "[SIN BACKUP] No se encontro ningun backup previo de '$ComputerName' en la share."
            Write-Warn "[SIN BACKUP] Borrar sin backup es extremadamente peligroso (sin recuperacion posible)."
        }
        Append-Output "" $script:White

        $confMsg = if ($prevBackup) {
            "Backup previo encontrado:`n  $($prevBackup.FullName)`n  Fecha: $($prevBackup.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))`n`n  ATENCION: Puede estar desactualizado.`n  No sustituye hacer un backup nuevo ahora."
        } else {
            "NO se encontro ningun backup previo para '$ComputerName'.`n  RIESGO MAXIMO: sin posibilidad de recuperacion."
        }

        if (-not (Confirm-Action (
            "!!! BORRADO SIN BACKUP !!!`n`n$confMsg`n`nPerfil a eliminar:`n  $profilePath en '$ComputerName'`n`nConfirmas el borrado SIN nuevo backup de '$usuario'?"
        ) "Perfilazo - BORRADO SIN BACKUP")) {
            Write-Info "Perfil NO borrado (cancelado)."; Write-Sep; Append-Output "" $script:White; return
        }

        Write-Warn "[SIN BACKUP] Borrando perfil '$usuario' SIN backup nuevo..."
        Set-Status "Borrando perfil '$usuario' [SIN BACKUP]..." ([System.Drawing.Color]::Orange)
        $delResult = Invoke-LocalOrRemote -ComputerName $ComputerName `
            -ArgumentList $profilePath -ScriptBlock {
                param([string]$localPath)
                try {
                    $prof = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
                                Where-Object { $_.LocalPath -eq $localPath }
                    if (-not $prof)   { return @{ Status='ERROR'; Details="Win32_UserProfile no encontrado para '$localPath'." } }
                    if ($prof.Loaded) { return @{ Status='WARN';  Details='Sesion activa. El perfil NO puede borrarse mientras esta cargado.' } }
                    Remove-CimInstance -InputObject $prof -ErrorAction Stop
                    return @{ Status='OK'; Details='Perfil eliminado correctamente.' }
                } catch { return @{ Status='ERROR'; Details=$_.Exception.Message } }
            }
        if (-not $delResult) { Write-Fail "[SIN BACKUP] Sin respuesta al borrar el perfil." }
        else {
            switch ($delResult.Status) {
                'OK'    { Write-Ok   "[SIN BACKUP] Perfil de '$usuario' eliminado. $($delResult.Details)" }
                'WARN'  { Write-Warn "[SIN BACKUP] Perfil NO borrado: $($delResult.Details)" }
                'ERROR' { Write-Fail "[SIN BACKUP] Error al borrar: $($delResult.Details)" }
            }
        }
        Write-Sep; Append-Output "" $script:White; return
    }

    # ── MODO SEGURO: backup nuevo + validacion + borrado ─────────────
    Write-Info "[CON BACKUP] Modo seguro activado."
    Append-Output "" $script:White

    # ── Verificar perfil ─────────────────────────────────────────────
    Write-Info "Verificando perfil: $profilePath"
    $existe = Invoke-LocalOrRemote -ComputerName $ComputerName -ArgumentList $profilePath -ScriptBlock {
        param([string]$p); return (Test-Path $p)
    }
    if (-not $existe) {
        Write-Fail "Perfil no encontrado: $profilePath en '$ComputerName'."
        Write-Sep; Append-Output "" $script:White; return
    }
    Write-Ok "Perfil encontrado: $profilePath"
    Append-Output "" $script:White

    # ── Calcular destino ─────────────────────────────────────────────
    $fecha   = (Get-Date).ToString('yyyy-MM-dd_HHmm')
    $destDir = "$destBase\${ComputerName}_${usuario}_$fecha"
    Write-Info "Destino: $destDir"
    Append-Output "" $script:White

    # ── Copia remota ─────────────────────────────────────────────────
    Write-Info "Copiando perfil..."
    Set-Status "Copiando perfil de '$usuario'..." ([System.Drawing.Color]::Yellow)
    $script:outputBox.Update()

    $copyResult = Invoke-LocalOrRemote -ComputerName $ComputerName `
        -ArgumentList $profilePath, $destDir, $rutaExtra `
        -OperationTimeoutMs 600000 `
        -ScriptBlock {
            param([string]$src, [string]$dst, [string]$extra)

            $folderDefs = @(
                @{ Name='Desktop';   Src="$src\Desktop"   },
                @{ Name='Downloads'; Src="$src\Downloads" },
                @{ Name='Documents'; Src="$src\Documents" },
                @{ Name='Favorites'; Src="$src\Favorites" },
                @{ Name='Pictures';  Src="$src\Pictures"  }
            )
            $fileDefs = @(
                @{ Name='Chrome_Bookmarks'; Src="$src\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"; Sub='BrowserBookmarks\Chrome' },
                @{ Name='Edge_Bookmarks';   Src="$src\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks";   Sub='BrowserBookmarks\Edge' }
            )

            $results = @()

            # Carpetas de perfil via robocopy
            foreach ($f in $folderDefs) {
                if (-not (Test-Path $f.Src)) {
                    $results += @{ Item=$f.Name; Status='SKIP'; Details='No existe' }; continue
                }
                $null = robocopy $f.Src "$dst\$($f.Name)" /E /COPY:DAT /R:1 /W:1 /NP /NJH /NJS /NS /NC /NFL /NDL 2>&1
                $ec = $LASTEXITCODE
                $st = if ($ec -ge 8) { 'ERROR' } else { 'OK' }
                $results += @{ Item=$f.Name; Status=$st; Details="rc=$ec" }
            }

            # Bookmarks via Copy-Item
            foreach ($fi in $fileDefs) {
                if (-not (Test-Path $fi.Src)) {
                    $results += @{ Item=$fi.Name; Status='SKIP'; Details='No existe' }; continue
                }
                $dstSub = "$dst\$($fi.Sub)"
                if (-not (Test-Path $dstSub)) { New-Item $dstSub -ItemType Directory -Force | Out-Null }
                try {
                    Copy-Item $fi.Src $dstSub -Force -ErrorAction Stop
                    $results += @{ Item=$fi.Name; Status='OK'; Details='Copiado' }
                } catch {
                    $results += @{ Item=$fi.Name; Status='ERROR'; Details=$_.Exception.Message }
                }
            }

            # Historial y configuracion de Cisco Jabber
            $jabberDefs = @(
                @{ Name='Jabber_History'; Src="$src\AppData\Local\Cisco\Unified Communications\Jabber\CSF\History"; Dst='Jabber\History' },
                @{ Name='Jabber_Config';  Src="$src\AppData\Roaming\Cisco\Unified Communications\Jabber";           Dst='Jabber\Config'  }
            )
            foreach ($j in $jabberDefs) {
                if (-not (Test-Path $j.Src)) {
                    $results += @{ Item=$j.Name; Status='WARN'; Details='Jabber no instalado en este perfil' }; continue
                }
                $null = robocopy $j.Src "$dst\$($j.Dst)" /E /COPY:DAT /R:1 /W:1 /NP /NJH /NJS /NS /NC /NFL /NDL 2>&1
                $ec = $LASTEXITCODE
                $st = if ($ec -ge 8) { 'ERROR' } else { 'OK' }
                $results += @{ Item=$j.Name; Status=$st; Details="rc=$ec" }
            }

            # Ruta extra opcional
            if (-not [string]::IsNullOrWhiteSpace($extra)) {
                if (-not (Test-Path $extra)) {
                    $results += @{ Item='RutaExtra'; Status='SKIP'; Details="No existe: $extra" }
                } elseif ((Get-Item $extra).PSIsContainer) {
                    $null = robocopy $extra "$dst\RutaExtra" /E /COPY:DAT /R:1 /W:1 /NP /NJH /NJS /NS /NC /NFL /NDL 2>&1
                    $ec = $LASTEXITCODE
                    $st = if ($ec -ge 8) { 'ERROR' } else { 'OK' }
                    $results += @{ Item='RutaExtra'; Status=$st; Details="rc=$ec" }
                } else {
                    $dstEx = "$dst\RutaExtra"
                    if (-not (Test-Path $dstEx)) { New-Item $dstEx -ItemType Directory -Force | Out-Null }
                    try {
                        Copy-Item $extra $dstEx -Force -ErrorAction Stop
                        $results += @{ Item='RutaExtra'; Status='OK'; Details='Copiado' }
                    } catch {
                        $results += @{ Item='RutaExtra'; Status='ERROR'; Details=$_.Exception.Message }
                    }
                }
            }

            return @{ Results=$results; BackupRoot=$dst }
        }

    if (-not $copyResult) {
        Write-Fail "Sin respuesta durante la copia. Abortando."
        Write-Sep; Append-Output "" $script:White; return
    }

    # ── Mostrar resultados de copia ───────────────────────────────────
    Write-Info "Resultado de copia:"
    $nOk = 0; $nErr = 0; $nSkip = 0
    foreach ($r in $copyResult.Results) {
        $color = switch ($r.Status) {
            'OK'    { [System.Drawing.Color]::LightGreen }
            'WARN'  { [System.Drawing.Color]::Yellow      }
            'SKIP'  { [System.Drawing.Color]::Gray        }
            default { [System.Drawing.Color]::Tomato      }
        }
        Append-Output ("  [{0,-6}] {1,-22}  {2}" -f $r.Status, $r.Item, $r.Details) $color
        switch ($r.Status) { 'OK' { $nOk++ } 'ERROR' { $nErr++ } default { $nSkip++ } }
    }
    Append-Output "" $script:White
    Write-Info "Resumen: Copiados=$nOk  Omitidos=$nSkip  Errores=$nErr"
    Append-Output "" $script:White

    # ── Validar backup ────────────────────────────────────────────────
    Write-Info "Validando backup..."
    if ($nOk -eq 0) {
        Write-Fail "No se copio ningun elemento. Perfil NO borrado."
        Write-Sep; Append-Output "" $script:White; return
    }

    $backupRoot = $copyResult.BackupRoot
    $validated = Invoke-LocalOrRemote -ComputerName $ComputerName `
        -ArgumentList $backupRoot -ScriptBlock {
            param([string]$root)
            if (-not (Test-Path $root)) { return $false }
            $firstFile = Get-ChildItem $root -Recurse -File -ErrorAction SilentlyContinue |
                         Select-Object -First 1
            return ($null -ne $firstFile)
        }

    if (-not $validated) {
        Write-Fail "Backup no verificado: carpeta vacia o inaccesible. Perfil NO borrado."
        Write-Sep; Append-Output "" $script:White; return
    }
    Write-Ok "Backup verificado: $backupRoot"
    if ($nErr -gt 0) {
        Write-Warn "$nErr elemento(s) con error en la copia. Revisar antes de proceder."
    }
    Append-Output "" $script:White

    # ── Confirmar borrado ─────────────────────────────────────────────
    $errWarn = if ($nErr -gt 0) { "`n`n  ATENCION: $nErr elemento(s) con error." } else { "" }
    if (-not (Confirm-Action (
        "Backup validado en:`n  $backupRoot`n`n" +
        "  Copiados=$nOk  Omitidos=$nSkip  Errores=$nErr$errWarn`n`n" +
        "Perfil a eliminar:`n  $profilePath en '$ComputerName'`n`n" +
        "¿Borrar el perfil de '$usuario'?"
    ) "Perfilazo - Confirmar borrado")) {
        Write-Info "Perfil NO borrado (cancelado)."
        Write-Sep; Append-Output "" $script:White; return
    }

    # ── Borrado de perfil via Win32_UserProfile ───────────────────────
    Write-Info "Borrando perfil via Win32_UserProfile..."
    Set-Status "Borrando perfil '$usuario'..." ([System.Drawing.Color]::Orange)

    $delResult = Invoke-LocalOrRemote -ComputerName $ComputerName `
        -ArgumentList $profilePath -ScriptBlock {
            param([string]$localPath)
            try {
                $prof = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop |
                            Where-Object { $_.LocalPath -eq $localPath }
                if (-not $prof) {
                    return @{ Status='ERROR'; Details="Win32_UserProfile no encontrado para '$localPath'." }
                }
                if ($prof.Loaded) {
                    return @{ Status='WARN'; Details='Sesion activa detectada. El perfil NO puede borrarse mientras esta cargado.' }
                }
                Remove-CimInstance -InputObject $prof -ErrorAction Stop
                return @{ Status='OK'; Details='Perfil eliminado correctamente.' }
            } catch {
                return @{ Status='ERROR'; Details=$_.Exception.Message }
            }
        }

    if (-not $delResult) {
        Write-Fail "Sin respuesta al borrar el perfil."
    } else {
        switch ($delResult.Status) {
            'OK'    { Write-Ok   "Perfil de '$usuario' eliminado. $($delResult.Details)" }
            'WARN'  { Write-Warn "Perfil NO borrado: $($delResult.Details)" }
            'ERROR' { Write-Fail "Error al borrar: $($delResult.Details)" }
        }
    }

    Write-Sep
    Append-Output "" $script:White
}

function Invoke-PerfilRestore {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Sep
    Write-Info "Restaurar Perfilazo en '$ComputerName'"
    Write-Sep
    Append-Output "" $script:White

    # ── Paso 1: Obtener y seleccionar perfil destino ──────────────────
    Write-Info "Obteniendo perfiles disponibles en '$ComputerName'..."
    $perfiles = Invoke-LocalOrRemote -ComputerName $ComputerName -ScriptBlock {
        $excluir = @('Public', 'Default', 'Default User', 'All Users')
        Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue |
            Where-Object {
                (-not $_.Special) -and
                ($_.LocalPath -like 'C:\Users\*') -and
                ((Split-Path $_.LocalPath -Leaf) -notin $excluir)
            } |
            Sort-Object LastUseTime -Descending |
            ForEach-Object {
                $lastUse = if ($_.LastUseTime) { $_.LastUseTime.ToString('yyyy-MM-dd HH:mm') } else { 'N/A' }
                @{ Name=(Split-Path $_.LocalPath -Leaf); LocalPath=$_.LocalPath; Loaded=$_.Loaded; LastUse=$lastUse }
            }
    }

    if (-not $perfiles) {
        Write-Fail "No hay perfiles disponibles en '$ComputerName'. El usuario debe haber iniciado sesion al menos una vez."
        Write-Sep; Append-Output "" $script:White; return
    }

    $idx = 0
    foreach ($p in @($perfiles)) {
        $idx++
        $loadColor = if ($p.Loaded) { [System.Drawing.Color]::Yellow } else { [System.Drawing.Color]::LightGreen }
        $loadStr   = if ($p.Loaded) { 'Loaded=True ' } else { 'Loaded=False' }
        Append-Output ("  [{0:D2}] {1,-20} | {2} | {3} | LastUse={4}" -f `
            $idx, $p.Name, $p.LocalPath, $loadStr, $p.LastUse) $loadColor
    }
    Append-Output "" $script:White

    $profilePath = Show-ProfilePicker -Profiles @($perfiles) `
                       -Title "Restaurar Perfilazo - '$ComputerName'"
    if (-not $profilePath) { Write-Warn "Cancelado."; return }
    $usuario = Split-Path $profilePath -Leaf

    # ── Paso 2: Ruta del backup origen ───────────────────────────────
    $backupRoot = (Get-Input "Ruta del backup a restaurar:" "Restaurar Perfilazo - Origen" "C:\Share").Trim()
    if ([string]::IsNullOrWhiteSpace($backupRoot)) { Write-Warn "Cancelado."; return }

    # ── Paso 3: Verificar perfil destino (sesion previa confirmada) ───
    Write-Info "Verificando perfil destino: $profilePath"
    $perfilOk = Invoke-LocalOrRemote -ComputerName $ComputerName -ArgumentList $profilePath -ScriptBlock {
        param([string]$p)
        if (-not (Test-Path $p)) { return $false }
        $prof = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue |
                    Where-Object { $_.LocalPath -eq $p }
        return ($null -ne $prof)
    }
    if (-not $perfilOk) {
        Write-Fail "Perfil '$profilePath' no confirmado en Win32_UserProfile. El usuario debe iniciar sesion al menos una vez antes de restaurar."
        Write-Sep; Append-Output "" $script:White; return
    }
    Write-Ok "Perfil destino confirmado: $profilePath"

    # ── Paso 4: Verificar y resolver ruta de backup ───────────────────
    Write-Info "Buscando backup en: $backupRoot"
    $backupResolve = Invoke-LocalOrRemote -ComputerName $ComputerName `
        -ArgumentList $backupRoot, $usuario, $ComputerName `
        -ScriptBlock {
            param([string]$root, [string]$user, [string]$pc)

            if (-not (Test-Path $root)) {
                return @{ OK=$false; Msg="Ruta no existe: $root"; Root=$root; Count=0 }
            }
            # La ruta ya apunta directamente al backup si tiene contenido esperado
            $tieneContenido = (Test-Path "$root\Desktop") -or
                              (Test-Path "$root\BrowserBookmarks") -or
                              (Test-Path "$root\Documents")
            if ($tieneContenido) {
                return @{ OK=$true; Root=$root; Count=1 }
            }
            # Buscar subcarpeta con el patron exacto generado por Perfilazo: PC_usuario_fecha
            $subs = @(Get-ChildItem $root -Directory -Filter "${pc}_${user}_*" `
                          -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
            if ($subs.Count -eq 0) {
                # Fallback: subcarpeta que contenga el nombre de usuario
                $subs = @(Get-ChildItem $root -Directory -ErrorAction SilentlyContinue |
                              Where-Object { $_.Name -like "*${user}*" } |
                              Sort-Object Name -Descending)
            }
            $cnt = $subs.Count
            if ($cnt -eq 0) {
                return @{ OK=$false; Msg="Sin estructura de backup ni subcarpetas con '$user' en: $root"; Root=$root; Count=0 }
            }
            $best = $subs[0].FullName
            return @{ OK=$true; Root=$best; Count=$cnt }
        }

    if (-not $backupResolve -or -not $backupResolve.OK) {
        $motivo = if ($backupResolve) { $backupResolve.Msg } else { "Sin respuesta del equipo." }
        Write-Fail "Backup no localizado: $motivo"
        Write-Sep; Append-Output "" $script:White; return
    }

    $backupRoot = $backupResolve.Root   # ruta real del backup (con timestamp)
    if ($backupResolve.Count -gt 1) {
        Write-Warn "Se encontraron $($backupResolve.Count) backups para '$usuario'. Usando el mas reciente."
    }
    Write-Ok "Backup localizado: $backupRoot"
    Append-Output "" $script:White

    # ── Paso 5: Confirmar ─────────────────────────────────────────────
    if (-not (Confirm-Action (
        "Restaurar backup:`n  $backupRoot`n`n" +
        "Perfil destino:`n  $profilePath en '$ComputerName'`n`n" +
        "Se restaurara: Desktop, Downloads, Documents, Favorites, Pictures,`n" +
        "Bookmarks y RutaExtra si existen en el backup.`n`n" +
        "El contenido existente NO se borra (se copia encima).`n`n" +
        "¿Continuar?"
    ) "Restaurar Perfilazo - Confirmar")) {
        Write-Info "Restauracion cancelada."
        Write-Sep; Append-Output "" $script:White; return
    }

    # ── Paso 6: Restaurar ─────────────────────────────────────────────
    Write-Info "Restaurando perfil de '$usuario'..."
    Set-Status "Restaurando perfil '$usuario'..." ([System.Drawing.Color]::Yellow)
    $script:outputBox.Update()

    $restResult = Invoke-LocalOrRemote -ComputerName $ComputerName `
        -ArgumentList $backupRoot, $profilePath `
        -OperationTimeoutMs 600000 `
        -ScriptBlock {
            param([string]$src, [string]$dst)

            $folderMap = @(
                @{ Name='Desktop';   Src="$src\Desktop";   Dst="$dst\Desktop"   },
                @{ Name='Downloads'; Src="$src\Downloads"; Dst="$dst\Downloads" },
                @{ Name='Documents'; Src="$src\Documents"; Dst="$dst\Documents" },
                @{ Name='Favorites'; Src="$src\Favorites"; Dst="$dst\Favorites" },
                @{ Name='Pictures';  Src="$src\Pictures";  Dst="$dst\Pictures"  }
            )
            $fileMap = @(
                @{ Name='Chrome_Bookmarks'
                   Src="$src\BrowserBookmarks\Chrome\Bookmarks"
                   Dst="$dst\AppData\Local\Google\Chrome\User Data\Default" },
                @{ Name='Edge_Bookmarks'
                   Src="$src\BrowserBookmarks\Edge\Bookmarks"
                   Dst="$dst\AppData\Local\Microsoft\Edge\User Data\Default" }
            )

            $results = @()

            foreach ($f in $folderMap) {
                if (-not (Test-Path $f.Src)) {
                    $results += @{ Item=$f.Name; Status='SKIP'; Details="No existe: $($f.Src)" }; continue
                }
                $null = robocopy $f.Src $f.Dst /E /COPY:DAT /R:1 /W:1 /NP /NJH /NJS /NS /NC /NFL /NDL 2>&1
                $ec = $LASTEXITCODE
                $st = if ($ec -ge 8) { 'ERROR' } else { 'OK' }
                $results += @{ Item=$f.Name; Status=$st; Details="rc=$ec" }
            }

            foreach ($fi in $fileMap) {
                if (-not (Test-Path $fi.Src)) {
                    $results += @{ Item=$fi.Name; Status='SKIP'; Details="No existe: $($fi.Src)" }; continue
                }
                if (-not (Test-Path $fi.Dst)) { New-Item $fi.Dst -ItemType Directory -Force | Out-Null }
                try {
                    Copy-Item $fi.Src $fi.Dst -Force -ErrorAction Stop
                    $results += @{ Item=$fi.Name; Status='OK'; Details='Restaurado' }
                } catch {
                    $results += @{ Item=$fi.Name; Status='ERROR'; Details=$_.Exception.Message }
                }
            }

            $extraSrc = "$src\RutaExtra"
            if (Test-Path $extraSrc) {
                $extraDst = "$dst\RutaExtra"
                $isDir = (Get-Item $extraSrc).PSIsContainer
                if ($isDir) {
                    $null = robocopy $extraSrc $extraDst /E /COPY:DAT /R:1 /W:1 /NP /NJH /NJS /NS /NC /NFL /NDL 2>&1
                    $ec = $LASTEXITCODE
                    $st = if ($ec -ge 8) { 'ERROR' } else { 'OK' }
                    $results += @{ Item='RutaExtra'; Status=$st; Details="rc=$ec (en $extraDst)" }
                } else {
                    if (-not (Test-Path $extraDst)) { New-Item $extraDst -ItemType Directory -Force | Out-Null }
                    try {
                        Copy-Item $extraSrc $extraDst -Force -ErrorAction Stop
                        $results += @{ Item='RutaExtra'; Status='OK'; Details="Restaurado en $extraDst" }
                    } catch {
                        $results += @{ Item='RutaExtra'; Status='ERROR'; Details=$_.Exception.Message }
                    }
                }
            }

            return @{ Results=$results }
        }

    if (-not $restResult) {
        Write-Fail "Sin respuesta durante la restauracion. Abortando."
        Write-Sep; Append-Output "" $script:White; return
    }

    Write-Info "Resultado de restauracion:"
    $nOk = 0; $nErr = 0; $nSkip = 0
    foreach ($r in $restResult.Results) {
        $color = switch ($r.Status) {
            'OK'    { [System.Drawing.Color]::LightGreen }
            'SKIP'  { [System.Drawing.Color]::Gray       }
            default { [System.Drawing.Color]::Tomato     }
        }
        Append-Output ("  [{0,-6}] {1,-22}  {2}" -f $r.Status, $r.Item, $r.Details) $color
        switch ($r.Status) { 'OK' { $nOk++ } 'ERROR' { $nErr++ } default { $nSkip++ } }
    }
    Append-Output "" $script:White
    Write-Info "Resumen: Restaurados=$nOk  Omitidos=$nSkip  Errores=$nErr"
    Append-Output "" $script:White

    if ($nOk -eq 0) {
        Write-Fail "No se restauro ningun elemento. Verifica la ruta del backup."
    } else {
        $st = if ($nErr -gt 0) { "con $nErr error(es)" } else { "correctamente" }
        Write-Ok "Restauracion completada $st. Restaurados=$nOk Omitidos=$nSkip"
    }

    Write-Sep
    Append-Output "" $script:White
}

function Invoke-CorporateCleanup {
    param([Parameter(Mandatory)][string]$ComputerName)

    Write-Sep
    Write-Info "Limpieza corporate en '$ComputerName'"
    Write-Sep
    Append-Output "" $script:White

    if (-not (Confirm-Action (
        "Se ejecutara limpieza de caches y temporales corporate en '$ComputerName':`n`n" +
        "  Windows Temp, Temp de usuarios, Chrome/Edge cache,`n" +
        "  Teams cache, Papelera`n`n" +
        "NO se borran documentos, descargas ni datos del usuario.`n`n" +
        "¿Continuar?"
    ) "Limpieza corporate - Confirmar")) {
        Write-Info "Limpieza cancelada."
        Write-Sep; Append-Output "" $script:White; return
    }

    Write-Info "Ejecutando limpieza..."
    Set-Status "Limpieza corporate '$ComputerName'..." ([System.Drawing.Color]::Yellow)
    $script:outputBox.Update()

    $cleanResult = Invoke-LocalOrRemote -ComputerName $ComputerName `
        -OperationTimeoutMs 300000 `
        -ScriptBlock {

            function Measure-FolderBytes {
                param([string]$Path)
                if (-not (Test-Path $Path)) { return [long]0 }
                $s = (Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue |
                          Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -eq $s) { return [long]0 }
                return [long]$s
            }

            function Clear-DirContents {
                param([string]$Path)
                if (-not (Test-Path $Path)) { return }
                Get-ChildItem $Path -Force -ErrorAction SilentlyContinue |
                    ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
            }

            $excluir = @('Public', 'Default', 'Default User', 'All Users')
            $usuarios = @(Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue |
                              Where-Object { $_.Name -notin $excluir })

            $items = @()

            # ── Windows Temp ──────────────────────────────────────────────
            $p = 'C:\Windows\Temp'
            $b = Measure-FolderBytes $p
            Clear-DirContents $p
            $freed = [Math]::Max([long]0, $b - (Measure-FolderBytes $p))
            $items += @{ Label='Windows Temp'; FreedBytes=$freed; Status='OK'; Details='' }

            # ── Temp de usuarios ──────────────────────────────────────────
            $utFreed = [long]0
            foreach ($u in $usuarios) {
                $tp = "$($u.FullName)\AppData\Local\Temp"
                if (Test-Path $tp) {
                    $b = Measure-FolderBytes $tp
                    Clear-DirContents $tp
                    $utFreed += [Math]::Max([long]0, $b - (Measure-FolderBytes $tp))
                }
            }
            $items += @{ Label='Temp usuarios'; FreedBytes=$utFreed; Status='OK'; Details='' }

            # ── Chrome/Edge cache ─────────────────────────────────────────
            $brFreed = [long]0
            $brSubs = @('Cache', 'Code Cache', 'GPUCache')
            foreach ($u in $usuarios) {
                foreach ($sub in $brSubs) {
                    foreach ($brPath in @(
                        "$($u.FullName)\AppData\Local\Google\Chrome\User Data\Default\$sub",
                        "$($u.FullName)\AppData\Local\Microsoft\Edge\User Data\Default\$sub"
                    )) {
                        if (Test-Path $brPath) {
                            $b = Measure-FolderBytes $brPath
                            Clear-DirContents $brPath
                            $brFreed += [Math]::Max([long]0, $b - (Measure-FolderBytes $brPath))
                        }
                    }
                }
            }
            $items += @{ Label='Chrome/Edge cache'; FreedBytes=$brFreed; Status='OK'; Details='' }

            # ── Teams cache ───────────────────────────────────────────────
            $tmFreed = [long]0
            $tmSubs = @('Cache', 'GPUCache', 'blob_storage', 'databases', 'IndexedDB')
            foreach ($u in $usuarios) {
                $tmRoots = @(
                    "$($u.FullName)\AppData\Local\Microsoft\Teams\current",
                    "$($u.FullName)\AppData\Roaming\Microsoft\Teams"
                )
                foreach ($tmRoot in $tmRoots) {
                    foreach ($sub in $tmSubs) {
                        $tp = "$tmRoot\$sub"
                        if (Test-Path $tp) {
                            $b = Measure-FolderBytes $tp
                            Clear-DirContents $tp
                            $tmFreed += [Math]::Max([long]0, $b - (Measure-FolderBytes $tp))
                        }
                    }
                }
            }
            $items += @{ Label='Teams cache'; FreedBytes=$tmFreed; Status='OK'; Details='' }

            # ── Papelera ──────────────────────────────────────────────────
            $rbPath = 'C:\$Recycle.Bin'
            $rbBefore = [long]0
            if (Test-Path $rbPath) {
                $rbBefore = (Get-ChildItem $rbPath -Recurse -Force -ErrorAction SilentlyContinue |
                                 Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -eq $rbBefore) { $rbBefore = [long]0 }
                Get-ChildItem $rbPath -Force -ErrorAction SilentlyContinue |
                    ForEach-Object { Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
                $rbAfter = (Get-ChildItem $rbPath -Recurse -Force -ErrorAction SilentlyContinue |
                                Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -eq $rbAfter) { $rbAfter = [long]0 }
                $rbFreed = [Math]::Max([long]0, [long]$rbBefore - [long]$rbAfter)
                $items += @{ Label='Papelera'; FreedBytes=$rbFreed; Status='OK'; Details='' }
            } else {
                $items += @{ Label='Papelera'; FreedBytes=[long]0; Status='SKIP'; Details='C:\$Recycle.Bin no encontrada' }
            }

            return @{ Items=$items }
        }

    if (-not $cleanResult) {
        Write-Fail "Sin respuesta durante la limpieza. Abortando."
        Write-Sep; Append-Output "" $script:White; return
    }

    Write-Info "Resultado limpieza corporate:"
    $totalFreed = [long]0
    foreach ($r in $cleanResult.Items) {
        $freed = [long]$r.FreedBytes

        $color = switch ($r.Status) {
            'OK'    { [System.Drawing.Color]::LightGreen }
            'SKIP'  { [System.Drawing.Color]::Gray       }
            default { [System.Drawing.Color]::Tomato     }
        }

        $freedStr = if ($freed -ge 1GB) { "{0:N2} GB" -f ($freed / 1GB) }
                    elseif ($freed -ge 1MB) { "{0:N1} MB" -f ($freed / 1MB) }
                    elseif ($freed -ge 1KB) { "{0:N0} KB" -f ($freed / 1KB) }
                    else { "$freed B" }

        $det = if ($r.Details) { "  ($($r.Details))" } else { '' }
        Append-Output ("  [{0,-5}] {1,-20}  {2} liberados{3}" -f `
            $r.Status, $r.Label, $freedStr, $det) $color
        $totalFreed += $freed
    }
    Append-Output "" $script:White

    $totalStr = if ($totalFreed -ge 1GB) { "{0:N2} GB" -f ($totalFreed / 1GB) }
                elseif ($totalFreed -ge 1MB) { "{0:N1} MB" -f ($totalFreed / 1MB) }
                else { "{0:N0} KB" -f ($totalFreed / 1KB) }
    Write-Ok "Total liberado: $totalStr"

    Write-Sep
    Append-Output "" $script:White
}

