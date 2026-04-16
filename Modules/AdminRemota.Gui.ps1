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

#region ═══════════════════════════════════════════════════════════
# VENTANA SECUNDARIA: NAC REMEDIATION
#═══════════════════════════════════════════════════════════════════

function Show-NacRemediationForm {

    # ── Rutas LDAP por servidor ───────────────────────────────────
    $nacPaths = @{
        "Getafe"  = "LDAP://nacget.intra.casa.corp:636/CN=NAC,DC=ds,DC=corp"
        "Sevilla" = "LDAP://nactab.intra.casa.corp:636/CN=NAC,DC=ds,DC=corp"
    }
    $nacRemPaths = @{
        "Getafe"  = "LDAP://nacget.intra.casa.corp:636/CN=Remediation,CN=NAC,DC=ds,DC=corp"
        "Sevilla" = "LDAP://nactab.intra.casa.corp:636/CN=Remediation,CN=NAC,DC=ds,DC=corp"
    }

    # ── Colores reutilizados ──────────────────────────────────────
    $cBgDark   = [System.Drawing.Color]::FromArgb(28,  28,  28)
    $cBgPanel  = [System.Drawing.Color]::FromArgb(45,  45,  48)
    $cBgOut    = [System.Drawing.Color]::FromArgb(16,  16,  16)
    $cWhite    = [System.Drawing.Color]::White
    $cSilver   = [System.Drawing.Color]::Silver
    $cAccent   = [System.Drawing.Color]::FromArgb(0,   122, 204)
    $cBtnGray  = [System.Drawing.Color]::FromArgb(62,   62,  66)
    $cBtnGreen = [System.Drawing.Color]::FromArgb(0,   130,  60)
    $cBtnPurp  = [System.Drawing.Color]::FromArgb(80,    0, 120)

    $fSmall = New-Object System.Drawing.Font("Segoe UI",  9)
    $fMono  = New-Object System.Drawing.Font("Consolas",  9)
    $fTitle = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

    # ── Helpers internos ─────────────────────────────────────────
    # AppendColor: anadir texto coloreado al RichTextBox de salida NAC
    # Helpers de escritura: usan $global:nacRtb para ser accesibles desde
    # cualquier scope, incluidos modulos dinamicos de GetNewClosure (PS5.1).
    # ReadOnly se desactiva temporalmente porque en WinForms, SelectionColor
    # no se aplica sobre texto ya insertado cuando ReadOnly=true, dejando el
    # texto en el color de insercion por defecto (negro sobre fondo oscuro).
    $AppendColor = {
        param([string]$Text, [System.Drawing.Color]$Color)
        if (-not $global:nacRtb) { return }
        $global:nacRtb.ReadOnly = $false
        $ss = $global:nacRtb.TextLength
        $global:nacRtb.AppendText($Text)
        $sl = $global:nacRtb.SelectionStart - $ss + 1
        $global:nacRtb.Select($ss, $sl)
        $global:nacRtb.SelectionColor = $Color
        $global:nacRtb.AppendText("`r`n")
        $global:nacRtb.SelectionStart = $global:nacRtb.TextLength
        $global:nacRtb.ScrollToCaret()
        $global:nacRtb.ReadOnly = $true
        [System.Windows.Forms.Application]::DoEvents()
    }
    $ClearOutput = {
        if ($global:nacRtb) {
            $global:nacRtb.ReadOnly = $false
            $global:nacRtb.Clear()
            $global:nacRtb.ReadOnly = $true
        }
    }
    # Escapa caracteres especiales RFC 4515 antes de insertar valores en filtros LDAP.
    # Obligatorio para: \ * ( ) y el caracter nulo.
    $EscapeLdap = {
        param([string]$Value)
        $Value = $Value.Replace('\', '\5c')
        $Value = $Value.Replace('*', '\2a')
        $Value = $Value.Replace('(', '\28')
        $Value = $Value.Replace(')', '\29')
        $Value = $Value -replace "`0", '\00'
        return $Value
    }

    # ── Formulario NAC ───────────────────────────────────────────
    $nacForm                 = New-Object System.Windows.Forms.Form
    $nacForm.Text            = "NAC Remediation"
    $nacForm.Size            = New-Object System.Drawing.Size(820, 520)
    $nacForm.MinimumSize     = New-Object System.Drawing.Size(720, 460)
    $nacForm.BackColor       = $cBgDark
    $nacForm.ForeColor       = $cWhite
    $nacForm.Font            = $fSmall
    $nacForm.StartPosition   = "CenterParent"
    $nacForm.FormBorderStyle = "Sizable"

    # ── Panel superior de controles ───────────────────────────────
    $nacTop             = New-Object System.Windows.Forms.Panel
    $nacTop.Dock        = "Top"
    $nacTop.Height      = 160
    $nacTop.BackColor   = $cBgPanel
    $nacForm.Controls.Add($nacTop)

    # Titulo
    $nacLblTitle           = New-Object System.Windows.Forms.Label
    $nacLblTitle.Text      = "  NAC REMEDIATION"
    $nacLblTitle.Font      = $fTitle
    $nacLblTitle.ForeColor = [System.Drawing.Color]::FromArgb(180, 100, 255)
    $nacLblTitle.AutoSize  = $true
    $nacLblTitle.Location  = New-Object System.Drawing.Point(8, 8)
    $nacTop.Controls.Add($nacLblTitle)

    # Servidor label + ComboBox
    $nacLblSrv          = New-Object System.Windows.Forms.Label
    $nacLblSrv.Text     = "Servidor:"
    $nacLblSrv.Location = New-Object System.Drawing.Point(10, 40)
    $nacLblSrv.Size     = New-Object System.Drawing.Size(65, 22)
    $nacLblSrv.TextAlign = "MiddleLeft"
    $nacTop.Controls.Add($nacLblSrv)

    $nacCboServer          = New-Object System.Windows.Forms.ComboBox
    $nacCboServer.Location = New-Object System.Drawing.Point(80, 40)
    $nacCboServer.Size     = New-Object System.Drawing.Size(110, 22)
    $nacCboServer.DropDownStyle = "DropDownList"
    $nacCboServer.BackColor    = [System.Drawing.Color]::FromArgb(55, 55, 60)
    $nacCboServer.ForeColor    = $cWhite
    [void]$nacCboServer.Items.Add("Getafe")
    [void]$nacCboServer.Items.Add("Sevilla")
    $nacCboServer.SelectedIndex = 0
    $nacTop.Controls.Add($nacCboServer)

    # Fila MAC
    $nacLblMac          = New-Object System.Windows.Forms.Label
    $nacLblMac.Text     = "MAC:"
    $nacLblMac.Location = New-Object System.Drawing.Point(10, 75)
    $nacLblMac.Size     = New-Object System.Drawing.Size(65, 22)
    $nacLblMac.TextAlign = "MiddleLeft"
    $nacTop.Controls.Add($nacLblMac)

    $nacTxtMac          = New-Object System.Windows.Forms.TextBox
    $nacTxtMac.Location = New-Object System.Drawing.Point(80, 75)
    $nacTxtMac.Size     = New-Object System.Drawing.Size(200, 22)
    $nacTxtMac.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 60)
    $nacTxtMac.ForeColor = $cWhite
    $nacTxtMac.BorderStyle = "FixedSingle"
    $nacTop.Controls.Add($nacTxtMac)

    $nacBtnCheckMac          = New-Object System.Windows.Forms.Button
    $nacBtnCheckMac.Text     = "Check MAC"
    $nacBtnCheckMac.Location = New-Object System.Drawing.Point(290, 73)
    $nacBtnCheckMac.Size     = New-Object System.Drawing.Size(110, 26)
    $nacBtnCheckMac.BackColor = $cBtnGray
    $nacBtnCheckMac.ForeColor = $cWhite
    $nacBtnCheckMac.FlatStyle = "Flat"
    $nacBtnCheckMac.FlatAppearance.BorderSize  = 1
    $nacBtnCheckMac.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $nacBtnCheckMac.Cursor   = "Hand"
    $nacTop.Controls.Add($nacBtnCheckMac)

    # Fila CN
    $nacLblCn           = New-Object System.Windows.Forms.Label
    $nacLblCn.Text      = "CN:"
    $nacLblCn.Location  = New-Object System.Drawing.Point(10, 110)
    $nacLblCn.Size      = New-Object System.Drawing.Size(65, 22)
    $nacLblCn.TextAlign = "MiddleLeft"
    $nacTop.Controls.Add($nacLblCn)

    $nacTxtCn           = New-Object System.Windows.Forms.TextBox
    $nacTxtCn.Location  = New-Object System.Drawing.Point(80, 110)
    $nacTxtCn.Size      = New-Object System.Drawing.Size(200, 22)
    $nacTxtCn.BackColor = [System.Drawing.Color]::FromArgb(55, 55, 60)
    $nacTxtCn.ForeColor = $cWhite
    $nacTxtCn.BorderStyle = "FixedSingle"
    $nacTop.Controls.Add($nacTxtCn)

    $nacBtnCheckCn          = New-Object System.Windows.Forms.Button
    $nacBtnCheckCn.Text     = "Check CN"
    $nacBtnCheckCn.Location = New-Object System.Drawing.Point(290, 108)
    $nacBtnCheckCn.Size     = New-Object System.Drawing.Size(110, 26)
    $nacBtnCheckCn.BackColor = $cBtnGray
    $nacBtnCheckCn.ForeColor = $cWhite
    $nacBtnCheckCn.FlatStyle = "Flat"
    $nacBtnCheckCn.FlatAppearance.BorderSize  = 1
    $nacBtnCheckCn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $nacBtnCheckCn.Cursor   = "Hand"
    $nacTop.Controls.Add($nacBtnCheckCn)

    # Boton Add Device (fila MAC, a la derecha)
    $nacBtnAdd          = New-Object System.Windows.Forms.Button
    $nacBtnAdd.Text     = "Add Device"
    $nacBtnAdd.Location = New-Object System.Drawing.Point(415, 73)
    $nacBtnAdd.Size     = New-Object System.Drawing.Size(110, 26)
    $nacBtnAdd.BackColor = $cBtnPurp
    $nacBtnAdd.ForeColor = $cWhite
    $nacBtnAdd.FlatStyle = "Flat"
    $nacBtnAdd.FlatAppearance.BorderSize  = 1
    $nacBtnAdd.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(100, 60, 140)
    $nacBtnAdd.Cursor   = "Hand"
    $nacTop.Controls.Add($nacBtnAdd)

    # Boton Limpiar salida
    $nacBtnClear          = New-Object System.Windows.Forms.Button
    $nacBtnClear.Text     = "Limpiar"
    $nacBtnClear.Location = New-Object System.Drawing.Point(415, 108)
    $nacBtnClear.Size     = New-Object System.Drawing.Size(110, 26)
    $nacBtnClear.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 66)
    $nacBtnClear.ForeColor = $cWhite
    $nacBtnClear.FlatStyle = "Flat"
    $nacBtnClear.FlatAppearance.BorderSize  = 1
    $nacBtnClear.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
    $nacBtnClear.Cursor   = "Hand"
    $nacTop.Controls.Add($nacBtnClear)

    # ── Area de salida ────────────────────────────────────────────
    # Se usa un panel intermedio (Dock=Fill) para garantizar que el RTB quede
    # debajo del panel superior, independientemente del orden de procesamiento
    # de Dock en PS5.1 WinForms. Sin el panel, Dock=Fill puede solaparse con
    # Dock=Top y ocultar el texto escrito en la zona superior del RTB.
    $nacOutPanel            = New-Object System.Windows.Forms.Panel
    $nacOutPanel.Dock       = "Fill"
    $nacOutPanel.BackColor  = $cBgOut
    $nacForm.Controls.Add($nacOutPanel)

    $global:nacRtb             = New-Object System.Windows.Forms.RichTextBox
    $global:nacRtb.Dock        = "Fill"
    $global:nacRtb.BackColor   = $cBgOut
    $global:nacRtb.ForeColor   = $cWhite
    $global:nacRtb.Font        = $fMono
    $global:nacRtb.ReadOnly    = $true
    $global:nacRtb.BorderStyle = "None"
    $global:nacRtb.ScrollBars  = "Vertical"
    $global:nacRtb.WordWrap    = $false
    $nacOutPanel.Controls.Add($global:nacRtb)
    # BringToFront en el panel Fill para que el Dock=Top (nacTop) se resuelva primero.
    # Sin esto nacOutPanel toma todo el espacio desde y=0 solapando nacTop (mismo
    # patron que el form principal: outputBox.BringToFront()).
    $nacOutPanel.BringToFront()

    # ── Closures de escritura con referencia local al RTB ──────────
    # ── Helper: mostrar propiedades de un resultado ADLDS ─────────
    $ShowResult = ({
        param($entry, [string]$ldapPath)
        $props = @("cn","deviceType","deviceZone","networkAddress","deviceRemediationID",
                   "adminDisplayName","whenCreated","description","devicemodel","device8021xcapable")
        & $AppendColor "  Ruta LDAP : $ldapPath" $cSilver
        foreach ($p in $props) {
            $val = $entry.Properties[$p]
            if ($val -and $val.Count -gt 0) {
                & $AppendColor ("  {0,-28}: {1}" -f $p, ($val | Select-Object -First 1)) $cWhite
            }
        }
        # Indicar en que contenedor esta el objeto
        $dn = ""
        if ($entry.Properties["distinguishedName"] -and $entry.Properties["distinguishedName"].Count -gt 0) {
            $dn = $entry.Properties["distinguishedName"][0]
        } elseif ($entry.Path) {
            $dn = $entry.Path
        }
        if ($dn -match "Remediation") {
            & $AppendColor "  >>> ESTADO: EN REMEDIATION <<<" ([System.Drawing.Color]::Yellow)
        } elseif ($dn -match "Migration") {
            & $AppendColor "  >>> ESTADO: EN MIGRATION <<<" ([System.Drawing.Color]::Orange)
        } elseif ($dn -match "Exception") {
            & $AppendColor "  >>> ESTADO: EN EXCEPTION <<<" ([System.Drawing.Color]::Cyan)
        } else {
            & $AppendColor "  >>> ESTADO: registrado (NAC normal) <<<" ([System.Drawing.Color]::LightGreen)
        }
    }).GetNewClosure()

    # ── Logica Check MAC ─────────────────────────────────────────
    $nacBtnCheckMac.Add_Click(({
        & $ClearOutput
        & $AppendColor "Check MAC clicked" ([System.Drawing.Color]::White)
        $mac = $nacTxtMac.Text.Trim().ToLower()
        $srv = $nacCboServer.SelectedItem
        if ([string]::IsNullOrEmpty($mac)) {
            & $AppendColor "WARN: Introduce una direccion MAC." ([System.Drawing.Color]::Tomato)
            return
        }
        & $AppendColor "INFO: Buscando MAC '$mac' en servidor '$srv'..." ([System.Drawing.Color]::Cyan)
        try {
            $ldapPath = $nacPaths[$srv]
            $de = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
            $ds = New-Object System.DirectoryServices.DirectorySearcher($de)
            $macE = & $EscapeLdap $mac
            $ds.Filter = "(|(networkAddress=$macE)(deviceRemediationID=$macE))"
            $ds.PropertiesToLoad.AddRange(@("cn","deviceType","deviceZone","networkAddress",
                "deviceRemediationID","adminDisplayName","whenCreated","description",
                "devicemodel","device8021xcapable","distinguishedName"))
            $ds.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            $results = $ds.FindAll()
            if ($results.Count -eq 0) {
                & $AppendColor "WARN: MAC no encontrada en NAC ($srv)." ([System.Drawing.Color]::Yellow)
            } else {
                & $AppendColor "OK: Encontrados $($results.Count) resultado(s):" ([System.Drawing.Color]::LightGreen)
                foreach ($r in $results) {
                    & $AppendColor "" $cWhite
                    & $ShowResult $r $ldapPath
                }
            }
            $results.Dispose()
        } catch {
            & $AppendColor "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        }
    }).GetNewClosure())

    # ── Logica Check CN ──────────────────────────────────────────
    $nacBtnCheckCn.Add_Click(({
        & $ClearOutput
        & $AppendColor "Check CN clicked" ([System.Drawing.Color]::White)
        $cn = $nacTxtCn.Text.Trim()
        $srv = $nacCboServer.SelectedItem
        if ([string]::IsNullOrEmpty($cn)) {
            & $AppendColor "WARN: Introduce un CN." ([System.Drawing.Color]::Tomato)
            return
        }
        & $AppendColor "INFO: Buscando CN '$cn' en servidor '$srv'..." ([System.Drawing.Color]::Cyan)
        try {
            $ldapPath = $nacPaths[$srv]
            $de = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
            $ds = New-Object System.DirectoryServices.DirectorySearcher($de)
            $ds.Filter = "(cn=$(& $EscapeLdap $cn))"
            $ds.PropertiesToLoad.AddRange(@("cn","deviceType","deviceZone","networkAddress",
                "deviceRemediationID","adminDisplayName","whenCreated","description",
                "devicemodel","device8021xcapable","distinguishedName"))
            $ds.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
            $results = $ds.FindAll()
            if ($results.Count -eq 0) {
                & $AppendColor "WARN: CN '$cn' no encontrado en NAC ($srv)." ([System.Drawing.Color]::Yellow)
            } else {
                & $AppendColor "OK: Encontrados $($results.Count) resultado(s):" ([System.Drawing.Color]::LightGreen)
                foreach ($r in $results) {
                    & $AppendColor "" $cWhite
                    & $ShowResult $r $ldapPath
                }
            }
            $results.Dispose()
        } catch {
            & $AppendColor "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        }
    }).GetNewClosure())

    # ── Logica Add Device ────────────────────────────────────────
    $nacBtnAdd.Add_Click(({
        & $ClearOutput
        $mac = $nacTxtMac.Text.Trim().ToLower()
        $cn  = $nacTxtCn.Text.Trim()
        $srv = $nacCboServer.SelectedItem

        # Si la MAC esta vacia, intentar resolverla desde ADLDS por CN
        if ([string]::IsNullOrEmpty($mac)) {
            if ([string]::IsNullOrEmpty($cn)) {
                & $AppendColor "WARN: Introduce al menos MAC o CN." ([System.Drawing.Color]::Tomato)
                return
            }
            & $AppendColor "INFO: MAC vacia, buscando networkAddress por CN '$cn'..." ([System.Drawing.Color]::Cyan)
            try {
                $ldapPath = $nacPaths[$srv]
                $de = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
                $ds = New-Object System.DirectoryServices.DirectorySearcher($de)
                $ds.Filter = "(cn=$(& $EscapeLdap $cn))"
                $ds.PropertiesToLoad.Add("networkAddress") | Out-Null
                $ds.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
                $found = $ds.FindOne()
                if ($found -and $found.Properties["networkAddress"].Count -gt 0) {
                    $mac = $found.Properties["networkAddress"][0].ToString().ToLower()
                    $nacTxtMac.Text = $mac
                    & $AppendColor "OK: MAC encontrada en ADLDS: $mac" ([System.Drawing.Color]::LightGreen)
                } else {
                    & $AppendColor "WARN: No se encontro networkAddress para CN '$cn'." ([System.Drawing.Color]::Yellow)
                    return
                }
            } catch {
                & $AppendColor "ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
                return
            }
        }

        # Validar formato MAC (xx:xx:xx:xx:xx:xx)
        if ($mac -notmatch '^([0-9a-f]{2}:){5}([0-9a-f]{2})$') {
            & $AppendColor "WARN: Formato MAC invalido. Usa aa:bb:cc:dd:ee:ff" ([System.Drawing.Color]::Tomato)
            return
        }

        # Requerir CN para crear el objeto
        if ([string]::IsNullOrEmpty($cn)) {
            & $AppendColor "WARN: CN requerido para crear el dispositivo." ([System.Drawing.Color]::Tomato)
            return
        }

        & $AppendColor "INFO: Creando dispositivo CN=$cn MAC=$mac en Remediation ($srv)..." ([System.Drawing.Color]::Cyan)
        try {
            $remPath = $nacRemPaths[$srv]
            $objOU   = [ADSI]$remPath
            $newDev  = $objOU.Create("device", "CN=$cn")
            $newDev.Put("deviceZone",          "ES")
            $newDev.Put("deviceType",          "Remediation-PC")
            $newDev.Put("deviceRemediationID", $mac)
            $newDev.Put("networkAddress",      $mac)
            $newDev.Put("adminDisplayName",    "$env:USERDOMAIN\$env:USERNAME")
            $newDev.SetInfo()
            & $AppendColor "OK: Dispositivo creado correctamente en Remediation ($srv)." ([System.Drawing.Color]::LightGreen)
            & $AppendColor "    CN  : $cn" $cWhite
            & $AppendColor "    MAC : $mac" $cWhite
            & $AppendColor "    User: $env:USERDOMAIN\$env:USERNAME" $cSilver
        } catch {
            & $AppendColor "ERROR: No se pudo crear el dispositivo: $($_.Exception.Message)" ([System.Drawing.Color]::Tomato)
        }
    }).GetNewClosure())

    # ── Boton Limpiar ────────────────────────────────────────────
    $nacBtnClear.Add_Click(({ & $ClearOutput }).GetNewClosure())

    # ── Mensaje de bienvenida ─────────────────────────────────────
    # Los AppendColor se emiten en Add_Shown (no antes de ShowDialog) porque
    # el handle del RichTextBox no existe hasta que ShowDialog lo crea; llamar
    # AppendText sin handle es un no-op en .NET 4.x y el texto se pierde.
    $nacForm.Add_Shown(({
        & $AppendColor ">>> NAC output ready <<<" ([System.Drawing.Color]::Cyan)
        & $AppendColor "NAC Remediation - herramienta de gestion de dispositivos ADLDS" $cSilver
        & $AppendColor "Selecciona servidor, introduce MAC o CN y usa los botones." $cSilver
        & $AppendColor "" $cWhite
    }).GetNewClosure())

    # ── Mostrar ventana modal ─────────────────────────────────────
    [void]$nacForm.ShowDialog($form)

    # Limpiar referencia RTB al cerrar
    $global:nacRtb = $null
}

#endregion

function New-GroupPanel {
    param([string]$Title, [int]$X, [int]$W, [int]$H = 134)
    $gp             = New-Object System.Windows.Forms.Panel
    $gp.Location    = New-Object System.Drawing.Point($X, 4)
    $gp.Size        = New-Object System.Drawing.Size($W, $H)
    $gp.BackColor   = [System.Drawing.Color]::FromArgb(38, 38, 42)
    $gp.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = "  $Title"
    $lbl.Location  = New-Object System.Drawing.Point(0, 0)
    $lbl.Size      = New-Object System.Drawing.Size($W, 18)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(0, 190, 255)
    $lbl.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 32)
    $lbl.Font      = $fontSmall
    $lbl.TextAlign = "MiddleLeft"
    $gp.Controls.Add($lbl)
    return $gp
}

function Resolve-FreshIP {
    param([string]$Hostname)
    # Metodo 1: pila completa del OS
    try {
        $ip = [System.Net.Dns]::GetHostAddresses($Hostname) |
              Where-Object {
                  $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork -and
                  -not $_.ToString().StartsWith('127.')    -and
                  -not $_.ToString().StartsWith('169.254.')
              } |
              Select-Object -First 1
        if ($ip) { return $ip.ToString() }
    } catch {}
    # Metodo 2: Resolve-DnsName directo al servidor DNS de la interfaz activa
    try {
        $srv = (Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.ServerAddresses.Count -gt 0 } |
                Select-Object -First 1).ServerAddresses | Select-Object -First 1
        $q = if ($srv) {
            Resolve-DnsName -Name $Hostname -Type A -DnsOnly -Server $srv -ErrorAction SilentlyContinue
        } else {
            Resolve-DnsName -Name $Hostname -Type A -DnsOnly -ErrorAction SilentlyContinue
        }
        $ip2 = ($q | Where-Object { $_.Type -eq 'A' } | Select-Object -First 1).IPAddress
        if ($ip2) { return $ip2 }
    } catch {}
    return $null
}

function Add-EquipoToList {
    param([string]$Name)
    $item           = New-Object System.Windows.Forms.ListViewItem("...")
    $item.Tag       = $Name
    $item.ForeColor = [System.Drawing.Color]::Gray
    $null           = $item.SubItems.Add($Name)    # SubItems[1] = Equipo
    $mast           = if ($script:MastStatus.ContainsKey($Name)) { $script:MastStatus[$Name] } else { '-' }
    $null           = $item.SubItems.Add($mast)    # SubItems[2] = Mast.
    $null           = $script:lvEquipos.Items.Add($item)
    return $item
}

function Set-MastStatus {
    param([string]$Name, [string]$Status)   # 'OK' | 'Error' | 'Pendiente'
    $script:MastStatus[$Name] = $Status
    $item = $script:lvEquipos.Items | Where-Object { $_.Tag -eq $Name } | Select-Object -First 1
    if (-not $item) { return }
    # Garantizar que existe SubItems[2]
    while ($item.SubItems.Count -lt 3) { $null = $item.SubItems.Add('') }
    $item.SubItems[2].Text = $Status
    # Color de fila: Mast. tiene prioridad sobre conectividad para dar feedback visual claro
    $item.ForeColor = switch ($Status) {
        'OK'       { [System.Drawing.Color]::LightGreen }
        'Error'    { [System.Drawing.Color]::Tomato }
        'Pendiente'{ [System.Drawing.Color]::Yellow }
        default    { [System.Drawing.Color]::Gray }
    }
}

function Update-EquipoCard {
    param($item)
    $hostname = $item.Tag
    # 1. Resolver IP: solo para display y clasificacion VPN/CABLE
    $freshIP = Resolve-FreshIP $hostname
    # 2. Ping SIEMPRE por hostname, nunca por IP.
    #    Pingar directamente la IP puede fallar si hay firewall ICMP en ese segmento,
    #    mientras que el hostname usa el routing completo del OS (igual que Test-Connection interno).
    $online  = Test-Connection -ComputerName $hostname -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($online) {
        $tipo                  = if ($freshIP -and ($freshIP.StartsWith("10.142.") -or $freshIP.StartsWith("10.99."))) { "VPN" } else { "CABLE" }
        $ipStr                 = if ($freshIP) { $freshIP } else { "?" }
        $item.SubItems[0].Text = "ONLINE"
        $item.ForeColor        = [System.Drawing.Color]::LightGreen
        $item.ToolTipText      = "$hostname  |  $tipo  |  $ipStr"
    } elseif ($freshIP) {
        # DNS resolvio pero ICMP no responde
        $item.SubItems[0].Text = "PING_FAIL"
        $item.ForeColor        = [System.Drawing.Color]::Orange
        $item.ToolTipText      = "$hostname  |  PING_FAIL  |  $freshIP"
    } else {
        # No resuelve y no responde al ping
        $item.SubItems[0].Text = "DNS_FAIL"
        $item.ForeColor        = [System.Drawing.Color]::OrangeRed
        $item.ToolTipText      = "$hostname  |  no resuelve (DNS_FAIL)"
    }
    [System.Windows.Forms.Application]::DoEvents()
}

function Refresh-EquipoEstados {
    foreach ($item in @($script:lvEquipos.Items)) {
        if ($item.Tag) { Update-EquipoCard $item }
    }
}

function Save-EquipoList {
    try {
        $names = @($script:lvEquipos.Items | ForEach-Object { $_.Tag })
        $json  = if ($names.Count -gt 0) { ConvertTo-Json -InputObject $names -Compress } else { '[]' }
        Set-Content -Path $script:EquiposFile -Value $json -Encoding UTF8
    } catch { <# sin permisos de escritura: se ignora silenciosamente #> }
}

function Load-EquipoList {
    if (-not (Test-Path $script:EquiposFile)) { return }
    try {
        $raw  = Get-Content -Path $script:EquiposFile -Raw -Encoding UTF8
        $data = $raw | ConvertFrom-Json
        foreach ($name in @($data)) {
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $already = $script:lvEquipos.Items | Where-Object { $_.Tag -eq $name }
            if ($already) { continue }
            $null = Add-EquipoToList $name
        }
    } catch { <# archivo corrupto: se ignora silenciosamente #> }
}

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
    # Equipo local: no necesita Test-Connection ni WinRM
    if (Test-IsLocal $computer) {
        Write-Ok "Equipo '$computer' es la maquina local (ejecucion directa)."
        return $computer
    }
    Set-Status "Comprobando conectividad con '$computer'..." ([System.Drawing.Color]::Yellow)
    # Resolver IP: solo informativo. El ping va SIEMPRE por hostname, igual que el boton Ping,
    # para mantener consistencia y evitar falsos negativos por ICMP filtrado por IP en VPN.
    $freshIP    = Resolve-FreshIP $computer
    if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
        Write-Fail "El equipo '$computer' no responde a ping (ICMP bloqueado o apagado)."
        Set-Status "Equipo no accesible" ([System.Drawing.Color]::Tomato)
        return $null
    }
    Write-Ok "Equipo '$computer' online."
    return $computer
}

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

function Show-RobocopyForm {
    $computer = $txtEquipo.Text.Trim()

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Copia Remota - Robocopy"
    $dlg.Size            = New-Object System.Drawing.Size(530, 320)
    $dlg.StartPosition   = "CenterParent"
    $dlg.BackColor       = $bgPanel
    $dlg.ForeColor       = $script:White
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false

    $lblComp           = New-Object System.Windows.Forms.Label
    if ($computer) { $lblComp.Text = "Equipo: $computer"; $lblComp.ForeColor = [System.Drawing.Color]::FromArgb(0, 190, 255) }
    else           { $lblComp.Text = "ATENCION: ningun equipo seleccionado"; $lblComp.ForeColor = [System.Drawing.Color]::Tomato }
    $lblComp.Font      = $fontSmall
    $lblComp.AutoSize  = $true
    $lblComp.Location  = New-Object System.Drawing.Point(12, 12)
    $dlg.Controls.Add($lblComp)

    $lblSrc           = New-Object System.Windows.Forms.Label
    $lblSrc.Text      = "Origen:"
    $lblSrc.ForeColor = $script:Silver
    $lblSrc.Font      = $fontSmall
    $lblSrc.AutoSize  = $true
    $lblSrc.Location  = New-Object System.Drawing.Point(12, 50)
    $dlg.Controls.Add($lblSrc)

    $txtSrc             = New-Object System.Windows.Forms.TextBox
    $txtSrc.Location    = New-Object System.Drawing.Point(70, 46)
    $txtSrc.Size        = New-Object System.Drawing.Size(430, 24)
    $txtSrc.BackColor   = [System.Drawing.Color]::FromArgb(55, 55, 58)
    $txtSrc.ForeColor   = $script:White
    $txtSrc.BorderStyle = "FixedSingle"
    $txtSrc.Font        = $fontUI
    $dlg.Controls.Add($txtSrc)

    $lblDst           = New-Object System.Windows.Forms.Label
    $lblDst.Text      = "Destino:"
    $lblDst.ForeColor = $script:Silver
    $lblDst.Font      = $fontSmall
    $lblDst.AutoSize  = $true
    $lblDst.Location  = New-Object System.Drawing.Point(12, 84)
    $dlg.Controls.Add($lblDst)

    $txtDst             = New-Object System.Windows.Forms.TextBox
    $txtDst.Location    = New-Object System.Drawing.Point(70, 80)
    $txtDst.Size        = New-Object System.Drawing.Size(430, 24)
    $txtDst.BackColor   = [System.Drawing.Color]::FromArgb(55, 55, 58)
    $txtDst.ForeColor   = $script:White
    $txtDst.BorderStyle = "FixedSingle"
    $txtDst.Font        = $fontUI
    $dlg.Controls.Add($txtDst)

    $txtLog             = New-Object System.Windows.Forms.RichTextBox
    $txtLog.Location    = New-Object System.Drawing.Point(12, 116)
    $txtLog.Size        = New-Object System.Drawing.Size(488, 118)
    $txtLog.BackColor   = [System.Drawing.Color]::FromArgb(16, 16, 16)
    $txtLog.ForeColor   = $script:White
    $txtLog.Font        = $fontMono
    $txtLog.ReadOnly    = $true
    $txtLog.BorderStyle = "None"
    $txtLog.WordWrap    = $false
    $dlg.Controls.Add($txtLog)

    $btnRun   = New-FlatButton "  Ejecutar" 12  246 100 28 ([System.Drawing.Color]::FromArgb(0, 100, 80))
    $btnClose = New-FlatButton "  Cerrar"   400 246 100 28 $btnGray
    $dlg.Controls.Add($btnRun)
    $dlg.Controls.Add($btnClose)

    $btnClose.Add_Click({ $dlg.Close() })

    $btnRun.Add_Click({
        $comp = $txtEquipo.Text.Trim()
        $src  = $txtSrc.Text.Trim()
        $dst  = $txtDst.Text.Trim()
        $txtLog.Clear()

        $appendLog = {
            param([string]$Msg, [System.Drawing.Color]$Col)
            $txtLog.SelectionStart  = $txtLog.TextLength
            $txtLog.SelectionLength = 0
            $txtLog.SelectionColor  = $Col
            $txtLog.AppendText("$Msg`r`n")
            $txtLog.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ([string]::IsNullOrWhiteSpace($comp)) {
            & $appendLog "ERROR: No hay equipo seleccionado en la ventana principal." ([System.Drawing.Color]::Tomato); return
        }
        if ([string]::IsNullOrWhiteSpace($src)) {
            & $appendLog "ERROR: La ruta origen no puede estar vacia." ([System.Drawing.Color]::Tomato); return
        }
        if ([string]::IsNullOrWhiteSpace($dst)) {
            & $appendLog "ERROR: La ruta destino no puede estar vacia." ([System.Drawing.Color]::Tomato); return
        }

        & $appendLog "Equipo:  $comp" ([System.Drawing.Color]::FromArgb(0, 190, 255))
        & $appendLog "Origen:  $src"  ([System.Drawing.Color]::Cyan)
        & $appendLog "Destino: $dst"  ([System.Drawing.Color]::Cyan)
        & $appendLog ""               ([System.Drawing.Color]::White)

        $btnRun.Enabled = $false
        $lblComp.Text   = "Equipo: $comp  [ejecutando...]"
        [System.Windows.Forms.Application]::DoEvents()

        $r = Invoke-LocalOrRemote -ComputerName $comp `
            -ArgumentList $src, $dst `
            -ScriptBlock {
                param([string]$origen, [string]$destino)
                if (-not (Test-Path $origen)) {
                    return @{ ExitCode=-1; Output="Ruta origen no encontrada: $origen" }
                }
                $out = robocopy $origen $destino /E /R:1 /W:1 /NFL /NDL /NP 2>&1
                return @{ ExitCode=$LASTEXITCODE; Output=($out -join "`n") }
            }

        if (-not $r) {
            & $appendLog "ERROR: Sin respuesta del equipo remoto." ([System.Drawing.Color]::Tomato)
        } else {
            $ec = $r.ExitCode
            if      ($ec -lt 0) { & $appendLog "ERROR: $($r.Output)" ([System.Drawing.Color]::Tomato) }
            elseif  ($ec -ge 8) {
                & $appendLog "ERROR: Robocopy rc=$ec. Revisar rutas y permisos." ([System.Drawing.Color]::Tomato)
                if ($r.Output) { & $appendLog $r.Output.Trim() ([System.Drawing.Color]::Silver) }
            }
            elseif  ($ec -eq 0) { & $appendLog "OK: Sin cambios (rc=0). Destino ya actualizado."  ([System.Drawing.Color]::Yellow) }
            else                { & $appendLog "OK: Copia completada correctamente (rc=$ec)."     ([System.Drawing.Color]::LightGreen) }
        }

        $btnRun.Enabled = $true
        $lblComp.Text   = "Equipo: $comp"
    })

    $dlg.ShowDialog($form) | Out-Null
    $dlg.Dispose()
}

function Show-WinRSSession {
    $computer = $script:txtEquipo.Text.Trim()

    if ([string]::IsNullOrEmpty($computer)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Introduce el nombre del equipo remoto primero.",
            "Campo requerido",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    if (Test-IsLocal $computer) {
        [System.Windows.Forms.MessageBox]::Show(
            "WinRS no tiene sentido contra el equipo local.",
            "Equipo local",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    if (-not (Confirm-Action "Se abrira una shell CMD remota en '$computer' via WinRS.`n`nNecesitas permisos de administrador en el equipo remoto.`n`nContinuar?" "Shell remota - WinRS")) {
        return
    }

    Write-Sep
    Write-Info "Abriendo shell remota en '$computer' via WinRS..."
    try {
        # /k mantiene la ventana CMD abierta al salir de la sesion WinRS
        Start-Process "cmd.exe" -ArgumentList "/k winrs -r:$computer cmd" -ErrorAction Stop
        Write-Ok "Ventana WinRS abierta para '$computer'."
        Set-Status "WinRS abierto en '$computer'" ([System.Drawing.Color]::LightGreen)
    } catch {
        Write-Fail "No se pudo abrir WinRS: $($_.Exception.Message)"
        Set-Status "Error al abrir WinRS" ([System.Drawing.Color]::Tomato)
    }
    Write-Sep
    Append-Output "" $script:White
}
