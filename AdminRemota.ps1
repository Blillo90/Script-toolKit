#Requires -Version 5.1
<#
.SYNOPSIS
    Herramienta de administracion remota unificada v2.17.3 (GUI)
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
    2.17.3
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
# Timeout global de 30s para conexion y operaciones WinRM.
# Se aplica automaticamente a todos los Invoke-Command sin SessionOption explicito.
$script:RemoteSessionOpt = New-PSSessionOption -OpenTimeout 10000 -OperationTimeout 30000
$PSDefaultParameterValues['Invoke-Command:SessionOption'] = $script:RemoteSessionOpt
$PSDefaultParameterValues['Invoke-Command:ErrorAction']   = 'SilentlyContinue'
$PSDefaultParameterValues['Get-CimInstance:ErrorAction']  = 'SilentlyContinue'

#region ═══════════════════════════════════════════════════════════
# CONSTANTES Y ESTADO GLOBAL
#═══════════════════════════════════════════════════════════════════

$script:outputBox       = $null
$script:statusLabel     = $null
$script:progressBar     = $null
$script:progressLabel   = $null
$script:cancelRequested = $false
$script:Modo            = "Nacional"   # "Nacional" | "Divisional"
$script:Target          = ""
$script:StepResults     = New-Object System.Collections.Generic.List[object]
$script:MastStatus      = @{}   # Name -> 'Pendiente' | 'OK' | 'Error'

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


# ── Modulos (dot-sourced) ──────────────────────────────────────────
# Cada modulo se carga en el scope del script llamante, preservando $script:* variables.
$PSScriptRoot_ = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

. "$PSScriptRoot_\Modules\AdminRemota.Logging.ps1"

. "$PSScriptRoot_\Modules\AdminRemota.Steps.ps1"

. "$PSScriptRoot_\Modules\AdminRemota.Remote.ps1"

. "$PSScriptRoot_\Modules\AdminRemota.Master.ps1"

. "$PSScriptRoot_\Modules\AdminRemota.Sccm.ps1"

. "$PSScriptRoot_\Modules\AdminRemota.System.ps1"

. "$PSScriptRoot_\Modules\AdminRemota.Perfilazo.ps1"

. "$PSScriptRoot_\Modules\AdminRemota.Gui.ps1"

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
# ── Bloque superior principal ─────────────────────────────────────
$topPanel           = New-Object System.Windows.Forms.Panel
$topPanel.Dock      = "Top"
$topPanel.Height    = 124
$topPanel.BackColor = $bgPanel
$form.Controls.Add($topPanel)

# TableLayoutPanel: garantiza filas con altura fija, sin solapamientos por z-order
$tlpTop                 = New-Object System.Windows.Forms.TableLayoutPanel
$tlpTop.Dock            = "Fill"
$tlpTop.ColumnCount     = 1
$tlpTop.RowCount        = 5
$tlpTop.BackColor       = [System.Drawing.Color]::Transparent
$tlpTop.Margin          = New-Object System.Windows.Forms.Padding(0)
$tlpTop.Padding         = New-Object System.Windows.Forms.Padding(0)
$tlpTop.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$topPanel.Controls.Add($tlpTop)

# Columna unica al 100%
$tlpTop.ColumnStyles.Clear()
$_cs          = New-Object System.Windows.Forms.ColumnStyle
$_cs.SizeType = [System.Windows.Forms.SizeType]::Percent
$_cs.Width    = 100
[void]$tlpTop.ColumnStyles.Add($_cs)

# Filas: alturas fijas en pixeles
$tlpTop.RowStyles.Clear()
foreach ($_h in @(60, 30, 28, 2, 34)) {
    $_rs          = New-Object System.Windows.Forms.RowStyle
    $_rs.SizeType = [System.Windows.Forms.SizeType]::Absolute
    $_rs.Height   = $_h
    [void]$tlpTop.RowStyles.Add($_rs)
}

# ── Fila 0: cabecera (titulo + subtitulo) ─────────────────────────
$pTitle           = New-Object System.Windows.Forms.Panel
$pTitle.Dock      = "Fill"
$pTitle.BackColor = [System.Drawing.Color]::Transparent
$pTitle.Margin    = New-Object System.Windows.Forms.Padding(0)
$tlpTop.Controls.Add($pTitle, 0, 0)

$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Text      = "  ADMINISTRACION REMOTA"
$lblTitle.Font      = $fontTitle
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 190, 255)
$lblTitle.AutoSize  = $true
$lblTitle.Location  = New-Object System.Drawing.Point(8, 4)
$pTitle.Controls.Add($lblTitle)

$lblSub             = New-Object System.Windows.Forms.Label
$lblSub.Text        = "  Accenture / Airbus  |  PowerShell 5.1"
$lblSub.Font        = $fontSmall
$lblSub.ForeColor   = $silver
$lblSub.AutoSize    = $true
$lblSub.Location    = New-Object System.Drawing.Point(8, 36)
$pTitle.Controls.Add($lblSub)

# ── Fila 1: contexto - equipo + Ping + resultado ──────────────────
$pContext           = New-Object System.Windows.Forms.Panel
$pContext.Dock      = "Fill"
$pContext.BackColor = [System.Drawing.Color]::Transparent
$pContext.Margin    = New-Object System.Windows.Forms.Padding(0)
$tlpTop.Controls.Add($pContext, 0, 1)

$lblEquipo           = New-Object System.Windows.Forms.Label
$lblEquipo.Text      = "Equipo:"
$lblEquipo.ForeColor = $silver
$lblEquipo.AutoSize  = $true
$lblEquipo.Location  = New-Object System.Drawing.Point(10, 8)
$pContext.Controls.Add($lblEquipo)

$txtEquipo             = New-Object System.Windows.Forms.TextBox
$txtEquipo.Width       = 240
$txtEquipo.Location    = New-Object System.Drawing.Point(68, 4)
$txtEquipo.BackColor   = [System.Drawing.Color]::FromArgb(55, 55, 58)
$txtEquipo.ForeColor   = $white
$txtEquipo.BorderStyle = "FixedSingle"
$txtEquipo.Font        = $fontUI
$pContext.Controls.Add($txtEquipo)

$btnPing = New-FlatButton "Ping" 315 1 62 28 $btnGray
$pContext.Controls.Add($btnPing)

# Label de resultado: ONLINE|VPN|IP / PING_FAIL|IP / DNS_FAIL
$lblPingResult             = New-Object System.Windows.Forms.Label
$lblPingResult.Text        = ""
$lblPingResult.Location    = New-Object System.Drawing.Point(385, 8)
$lblPingResult.Size        = New-Object System.Drawing.Size(530, 18)
$lblPingResult.ForeColor   = $white
$lblPingResult.BackColor   = [System.Drawing.Color]::Transparent
$lblPingResult.Font        = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$pContext.Controls.Add($lblPingResult)

# Referencia compartida para closures del panel lateral
$script:EquipoInputBox = $txtEquipo

# ── Fila 2: opciones - Modo  |  NAC Remediation (derecha) ─────────
$pOptions           = New-Object System.Windows.Forms.Panel
$pOptions.Dock      = "Fill"
$pOptions.BackColor = [System.Drawing.Color]::Transparent
$pOptions.Margin    = New-Object System.Windows.Forms.Padding(0)
$tlpTop.Controls.Add($pOptions, 0, 2)

$lblModo           = New-Object System.Windows.Forms.Label
$lblModo.Text      = "Modo:"
$lblModo.ForeColor = $silver
$lblModo.AutoSize  = $true
$lblModo.Location  = New-Object System.Drawing.Point(10, 7)
$pOptions.Controls.Add($lblModo)

$rbNacional           = New-Object System.Windows.Forms.RadioButton
$rbNacional.Text      = "Nacional"
$rbNacional.ForeColor = $white
$rbNacional.AutoSize  = $true
$rbNacional.Checked   = $true
$rbNacional.Location  = New-Object System.Drawing.Point(58, 4)
$rbNacional.Font      = $fontSmall
$pOptions.Controls.Add($rbNacional)

$rbDivisional           = New-Object System.Windows.Forms.RadioButton
$rbDivisional.Text      = "Divisional"
$rbDivisional.ForeColor = $white
$rbDivisional.AutoSize  = $true
$rbDivisional.Location  = New-Object System.Drawing.Point(148, 4)
$rbDivisional.Font      = $fontSmall
$pOptions.Controls.Add($rbDivisional)

$rbNacional.Add_CheckedChanged({
    $script:Modo = if ($rbNacional.Checked) { "Nacional" } else { "Divisional" }
})


# ── Fila 3: separador visual ──────────────────────────────────────
$pSep           = New-Object System.Windows.Forms.Panel
$pSep.Dock      = "Fill"
$pSep.BackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$pSep.Margin    = New-Object System.Windows.Forms.Padding(0)
$tlpTop.Controls.Add($pSep, 0, 3)

# ── Fila 4: bloque de progreso + estado ───────────────────────────
# progressBar + Cancelar en la misma fila para asociarlos visualmente.
# Limpiar como utilidad secundaria al extremo derecho.
$pProg           = New-Object System.Windows.Forms.Panel
$pProg.Dock      = "Fill"
$pProg.BackColor = [System.Drawing.Color]::Transparent
$pProg.Margin    = New-Object System.Windows.Forms.Padding(0)
$tlpTop.Controls.Add($pProg, 0, 4)

$script:progressBar          = New-Object System.Windows.Forms.ProgressBar
$script:progressBar.Location = New-Object System.Drawing.Point(10, 8)
$script:progressBar.Size     = New-Object System.Drawing.Size(430, 18)
$script:progressBar.Minimum  = 0
$script:progressBar.Maximum  = 100
$script:progressBar.Value    = 0
$script:progressBar.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
$pProg.Controls.Add($script:progressBar)

$btnCancel         = New-FlatButton "  Cancelar" 446 3 100 28 ([System.Drawing.Color]::FromArgb(200, 100, 0))
$btnCancel.Enabled = $false
$pProg.Controls.Add($btnCancel)

$script:progressLabel           = New-Object System.Windows.Forms.Label
$script:progressLabel.Location  = New-Object System.Drawing.Point(552, 9)
$script:progressLabel.Size      = New-Object System.Drawing.Size(175, 18)
$script:progressLabel.ForeColor = $silver
$script:progressLabel.Font      = $fontSmall
$script:progressLabel.Text      = ""
$script:progressLabel.TextAlign = "MiddleLeft"
$pProg.Controls.Add($script:progressLabel)

$btnClear = New-FlatButton "  Limpiar" 735 3 80 28 $btnGray
$pProg.Controls.Add($btnClear)

# ── Panel de acciones agrupadas ───────────────────────────────────
#   5 grupos horizontales: Diagnostico | SCCM/Politicas | Sistema | Usuario | Sensibles
$actionPanel           = New-Object System.Windows.Forms.Panel
$actionPanel.Dock      = "Top"
$actionPanel.Height    = 144
$actionPanel.BackColor = $bgPanel
$form.Controls.Add($actionPanel)

# ── G1: Diagnostico (x=4, w=215) ─────────────────────────────────
$gDiag = New-GroupPanel "Diagnostico" 4 215
$bw1   = 205
$btnMaster     = New-FlatButton "  Masterizacion"     4  22 $bw1 24 $accent
$btnSoftware   = New-FlatButton "  Software SCCM"     4  50 $bw1 24 $accent
$btnInfo       = New-FlatButton "  Info del Sistema"  4  78 $bw1 24 ([System.Drawing.Color]::FromArgb(40, 110, 60))
$btnSccmRepair = New-FlatButton "  SCCM Repair"       4 106 $bw1 24 ([System.Drawing.Color]::FromArgb(0, 80, 100))
foreach ($b in @($btnMaster, $btnSoftware, $btnInfo, $btnSccmRepair)) { $gDiag.Controls.Add($b) }

# ── G2: SCCM / Politicas (x=223, w=165) ──────────────────────────
$gSccm         = New-GroupPanel "SCCM / Politicas" 223 165
$bw2           = 155
$btnGpUpdate   = New-FlatButton "  GPUpdate /force" 4 22 $bw2 24 ([System.Drawing.Color]::FromArgb(0, 90, 160))
$btnSccmCycles = New-FlatButton "  Ciclos SCCM"     4 50 $bw2 24 ([System.Drawing.Color]::FromArgb(0, 90, 160))
foreach ($b in @($btnGpUpdate, $btnSccmCycles)) { $gSccm.Controls.Add($b) }

# ── G3: Sistema (x=392, w=180) ───────────────────────────────────
$gSistema  = New-GroupPanel "Sistema" 392 180
$bw3       = 170
$btnRepair  = New-FlatButton "  DISM + SFC"          4 22 $bw3 24 ([System.Drawing.Color]::FromArgb(130, 80, 0))
$btnChkdsk  = New-FlatButton "  ChkDsk /r"           4 50 $bw3 24 ([System.Drawing.Color]::FromArgb(150, 60, 0))
$btnCleanup  = New-FlatButton "  Limpieza temporales" 4 78  $bw3 24 ([System.Drawing.Color]::FromArgb(0, 100, 80))
$btnRobocopy = New-FlatButton "  Copia remota"        4 106 $bw3 24 ([System.Drawing.Color]::FromArgb(0, 90, 130))
foreach ($b in @($btnRepair, $btnChkdsk, $btnCleanup, $btnRobocopy)) { $gSistema.Controls.Add($b) }

# ── G4: Usuario (x=576, w=170) ───────────────────────────────────
$gUsuario         = New-GroupPanel "Usuario" 576 170
$bw4              = 160
$btnPerfilazo     = New-FlatButton "  Perfilazo"           4 22 $bw4 24 $accent
$btnPerfilRestore = New-FlatButton "  Restaurar Perfilazo" 4 50 $bw4 24 ([System.Drawing.Color]::FromArgb(0, 80, 140))
foreach ($b in @($btnPerfilazo, $btnPerfilRestore)) { $gUsuario.Controls.Add($b) }

# ── G5: Sensibles (x=750, w=155) ─────────────────────────────────
$gSensibles = New-GroupPanel "Sensibles" 750 155
$bw5        = 145
$btnRestart = New-FlatButton "  Reiniciar"   4 22 $bw5 24 ([System.Drawing.Color]::FromArgb(160, 80, 0))
$btnUsb     = New-FlatButton "  Borrar USB"  4 50 $bw5 24 $btnRed
$btnUsb.Enabled = $false   # Temporalmente deshabilitado - pendiente reactivacion
$btnWinRS   = New-FlatButton "  Shell remota (WinRS)" 4 78 $bw5 24 ([System.Drawing.Color]::FromArgb(80, 30, 120))
foreach ($b in @($btnRestart, $btnUsb, $btnWinRS)) { $gSensibles.Controls.Add($b) }

foreach ($g in @($gDiag, $gSccm, $gSistema, $gUsuario, $gSensibles)) {
    $actionPanel.Controls.Add($g)
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
$statusBar           = New-Object System.Windows.Forms.Panel
$statusBar.Dock      = "Bottom"
$statusBar.Height    = 24
$statusBar.BackColor = $accent
$form.Controls.Add($statusBar)

# ── Panel de seguimiento de equipos ──────────────────────────────
# TableLayoutPanel (tlpRight) garantiza que la cabecera no solape la lista.
# z-order de WinForms con Dock=Top/Fill/Bottom es ambiguo en PS5.1;
# TableLayoutPanel asigna espacio real a cada fila de forma determinista.
$rightPanel             = New-Object System.Windows.Forms.Panel
$rightPanel.Dock        = "Right"
$rightPanel.Width       = 210
$rightPanel.BackColor   = [System.Drawing.Color]::FromArgb(38, 38, 42)
$form.Controls.Add($rightPanel)

$tlpRight                 = New-Object System.Windows.Forms.TableLayoutPanel
$tlpRight.Dock            = "Fill"
$tlpRight.ColumnCount     = 1
$tlpRight.RowCount        = 4
$tlpRight.BackColor       = [System.Drawing.Color]::Transparent
$tlpRight.Margin          = New-Object System.Windows.Forms.Padding(0)
$tlpRight.Padding         = New-Object System.Windows.Forms.Padding(0)
$tlpRight.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$rightPanel.Controls.Add($tlpRight)

$tlpRight.ColumnStyles.Clear()
$_rc          = New-Object System.Windows.Forms.ColumnStyle
$_rc.SizeType = [System.Windows.Forms.SizeType]::Percent
$_rc.Width    = 100
[void]$tlpRight.ColumnStyles.Add($_rc)

$tlpRight.RowStyles.Clear()
# Fila 0: NAC Remediation 32px
$_rr_nac          = New-Object System.Windows.Forms.RowStyle
$_rr_nac.SizeType = [System.Windows.Forms.SizeType]::Absolute
$_rr_nac.Height   = 32
[void]$tlpRight.RowStyles.Add($_rr_nac)
# Fila 1: cabecera fija 26px
$_rr0          = New-Object System.Windows.Forms.RowStyle
$_rr0.SizeType = [System.Windows.Forms.SizeType]::Absolute
$_rr0.Height   = 26
[void]$tlpRight.RowStyles.Add($_rr0)
# Fila 2: lista ocupa todo el espacio restante
$_rr1          = New-Object System.Windows.Forms.RowStyle
$_rr1.SizeType = [System.Windows.Forms.SizeType]::Percent
$_rr1.Height   = 100
[void]$tlpRight.RowStyles.Add($_rr1)
# Fila 3: toolbar inferior fija 96px
$_rr2          = New-Object System.Windows.Forms.RowStyle
$_rr2.SizeType = [System.Windows.Forms.SizeType]::Absolute
$_rr2.Height   = 96
[void]$tlpRight.RowStyles.Add($_rr2)

# Fila 0 del TLP: NAC Remediation (herramienta avanzada)
$nacRow           = New-Object System.Windows.Forms.Panel
$nacRow.Dock      = "Fill"
$nacRow.BackColor = [System.Drawing.Color]::FromArgb(38, 38, 42)
$nacRow.Margin    = New-Object System.Windows.Forms.Padding(0)
$btnNAC = New-FlatButton "  NAC Remediation" 3 3 202 26 ([System.Drawing.Color]::FromArgb(80, 0, 120))
$btnNAC.Add_Click({ Show-NacRemediationForm })
$nacRow.Controls.Add($btnNAC)
$tlpRight.Controls.Add($nacRow, 0, 0)

# Fila 1 del TLP: cabecera
$rightHeader             = New-Object System.Windows.Forms.Panel
$rightHeader.Dock        = "Fill"
$rightHeader.BackColor   = [System.Drawing.Color]::FromArgb(28, 28, 32)
$rightHeader.Margin      = New-Object System.Windows.Forms.Padding(0)
$tlpRight.Controls.Add($rightHeader, 0, 1)

$lblEquiposTitle           = New-Object System.Windows.Forms.Label
$lblEquiposTitle.Text      = "  Equipos en seguimiento"
$lblEquiposTitle.Dock      = "Fill"
$lblEquiposTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 190, 255)
$lblEquiposTitle.Font      = $fontSmall
$lblEquiposTitle.TextAlign = "MiddleLeft"
$rightHeader.Controls.Add($lblEquiposTitle)

# Fila 2 del TLP: ListView (llena el espacio restante entre cabecera y toolbar)
$script:lvEquipos               = New-Object System.Windows.Forms.ListView
$script:lvEquipos.Dock          = "Fill"
$script:lvEquipos.View          = [System.Windows.Forms.View]::Details
$script:lvEquipos.FullRowSelect = $true
$script:lvEquipos.MultiSelect   = $false
$script:lvEquipos.HideSelection = $false
$script:lvEquipos.HeaderStyle   = [System.Windows.Forms.ColumnHeaderStyle]::None
$script:lvEquipos.BackColor     = [System.Drawing.Color]::FromArgb(32, 32, 35)
$script:lvEquipos.ForeColor     = [System.Drawing.Color]::White
$script:lvEquipos.BorderStyle   = "None"
$script:lvEquipos.Font          = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$script:lvEquipos.ShowItemToolTips = $true
$script:lvEquipos.Margin        = New-Object System.Windows.Forms.Padding(0)
$null = $script:lvEquipos.Columns.Add("Estado", 56)
$null = $script:lvEquipos.Columns.Add("Equipo", 96)
$null = $script:lvEquipos.Columns.Add("Mast.", 54)
$tlpRight.Controls.Add($script:lvEquipos, 0, 2)

# Fila 3 del TLP: toolbar de gestion de equipos
$rightBtnPanel           = New-Object System.Windows.Forms.Panel
$rightBtnPanel.Dock      = "Fill"
$rightBtnPanel.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
$rightBtnPanel.Margin    = New-Object System.Windows.Forms.Padding(0)
$tlpRight.Controls.Add($rightBtnPanel, 0, 3)

$btnAddEquipo      = New-FlatButton "Anadir equipo"    3  3 202 26 ([System.Drawing.Color]::FromArgb(0, 100, 50))
$btnRemoveEquipo   = New-FlatButton "Quitar selec."    3 33 202 26 $btnGray
$btnRefreshEquipos = New-FlatButton "Refrescar estado" 3 63 202 26 ([System.Drawing.Color]::FromArgb(30, 60, 110))
$rightBtnPanel.Controls.Add($btnAddEquipo)
$rightBtnPanel.Controls.Add($btnRemoveEquipo)
$rightBtnPanel.Controls.Add($btnRefreshEquipos)

$script:outputBox.BringToFront()   # Fill ocupa el espacio restante tras Right/Bottom/Top

$script:statusLabel           = New-Object System.Windows.Forms.Label
$script:statusLabel.Text      = "  Listo - introduce un equipo y pulsa una accion"
$script:statusLabel.Dock      = "Fill"
$script:statusLabel.ForeColor = $white
$script:statusLabel.Font      = $fontSmall
$script:statusLabel.TextAlign = "MiddleLeft"
$statusBar.Controls.Add($script:statusLabel)

# ── Helpers del panel lateral de equipos ─────────────────────────

# Archivo de persistencia en el mismo directorio que el script
$_equiposDir        = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:EquiposFile = Join-Path $_equiposDir "equipos_seguimiento.json"

# ── Helpers del ListView de equipos ──────────────────────────────

# Resuelve la IP actual del hostname.
# Metodo 1: [System.Net.Dns]::GetHostAddresses - usa la pila completa del OS
#   (hosts, DNS, NetBIOS), los mismos mecanismos que Test-Connection.
# Metodo 2 (fallback): Resolve-DnsName directo al servidor DNS de la interfaz.
# Devuelve la primera IPv4 util (excluye loopback y APIPA) o $null.
# Seleccion en ListView -> carga equipo en textbox principal
$script:lvEquipos.Add_SelectedIndexChanged({
    if ($script:lvEquipos.SelectedItems.Count -eq 0) { return }
    $script:EquipoInputBox.Text = $script:lvEquipos.SelectedItems[0].Tag
})
# ── Helpers de control de UI ──────────────────────────────────────

# Lista de botones de accion (todos excepto Cancelar y Limpiar)
$script:ActionButtons = @(
    $btnMaster, $btnSoftware, $btnInfo, $btnSccmRepair,
    $btnGpUpdate, $btnSccmCycles,
    $btnRepair, $btnChkdsk, $btnCleanup,
    $btnPerfilazo, $btnPerfilRestore,
    $btnRestart, $btnPing, $btnRobocopy, $btnWinRS
)
# ── Eventos ───────────────────────────────────────────────────────

$btnPing.Add_Click({
    $computer = $txtEquipo.Text.Trim()
    if ([string]::IsNullOrEmpty($computer)) { return }

    $lblPingResult.Text      = "Comprobando..."
    $lblPingResult.ForeColor = [System.Drawing.Color]::Yellow
    Set-Status "Haciendo ping a '$computer'..." ([System.Drawing.Color]::Yellow)
    Write-Sep

    # 1. Resolver IP: solo para display y clasificacion VPN/CABLE
    $freshIP = Resolve-FreshIP $computer
    # 2. Ping SIEMPRE por hostname, nunca por IP.
    #    Pingar directamente la IP puede fallar si hay firewall ICMP en ese segmento
    #    mientras que el hostname usa el routing completo del OS.
    $online  = Test-Connection -ComputerName $computer -Count 1 -Quiet -ErrorAction SilentlyContinue

    if ($online) {
        $tipo  = if ($freshIP -and ($freshIP.StartsWith("10.142.") -or $freshIP.StartsWith("10.99."))) { "VPN" } else { "CABLE" }
        $ipStr = if ($freshIP) { $freshIP } else { "?" }
        $lblPingResult.Text      = "ONLINE  |  $tipo  |  $ipStr"
        $lblPingResult.ForeColor = [System.Drawing.Color]::LightGreen
        Set-Status "  '$computer' ONLINE  |  $tipo  |  $ipStr" ([System.Drawing.Color]::LightGreen)
        Write-Ok "Ping OK -> '$computer' | Tipo: $tipo | IP: $ipStr"
    } elseif ($freshIP) {
        $lblPingResult.Text      = "PING_FAIL  |  $freshIP"
        $lblPingResult.ForeColor = [System.Drawing.Color]::Orange
        Set-Status "  '$computer' PING_FAIL  |  $freshIP" ([System.Drawing.Color]::Orange)
        Write-Warn "Ping FAIL -> '$computer' resuelve a $freshIP pero no responde al ping."
    } else {
        $lblPingResult.Text      = "DNS_FAIL"
        $lblPingResult.ForeColor = [System.Drawing.Color]::Tomato
        Set-Status "  '$computer' DNS_FAIL (no resuelve y no responde)" ([System.Drawing.Color]::Tomato)
        Write-Fail "DNS FAIL -> '$computer' no resuelve y no responde al ping."
    }

    # Actualizar el item del ListView si el equipo esta en seguimiento
    $lvItem = $script:lvEquipos.Items | Where-Object { $_.Tag -eq $computer } | Select-Object -First 1
    if ($lvItem) {
        if ($online) {
            $lvItem.SubItems[0].Text = "ONLINE"
            $lvItem.ToolTipText      = "$computer  |  $tipo  |  $ipStr"
        } elseif ($freshIP) {
            $lvItem.SubItems[0].Text = "PING_FAIL"
            $lvItem.ToolTipText      = "$computer  |  PING_FAIL  |  $freshIP"
        } else {
            $lvItem.SubItems[0].Text = "DNS_FAIL"
            $lvItem.ToolTipText      = "$computer  |  no resuelve (DNS_FAIL)"
        }
        $mast = if ($script:MastStatus.ContainsKey($computer)) { $script:MastStatus[$computer] } else { $null }
        $lvItem.ForeColor = if     ($mast -eq 'OK')       { [System.Drawing.Color]::LightGreen }
                            elseif ($mast -eq 'Error')     { [System.Drawing.Color]::Tomato     }
                            elseif ($mast -eq 'Pendiente') { [System.Drawing.Color]::Yellow     }
                            elseif ($online)               { [System.Drawing.Color]::LightGreen }
                            elseif ($freshIP)              { [System.Drawing.Color]::Orange     }
                            else                           { [System.Drawing.Color]::OrangeRed  }
        $script:lvEquipos.Refresh()
    }

    Write-Sep
    Append-Output "" $white
})

$btnMaster.Add_Click({
    $target = Get-ValidComputer
    if (-not $target) { return }
    Set-MastStatus -Name $target -Status 'Pendiente'
    Invoke-ActionButton -ComputerName $target `
        -StatusMsg "Comprobando masterizacion de '$target'..." `
        -Action    { Invoke-MasterCheck -ComputerName $target }
    # Evaluar resultado con los mismos criterios que Show-Summary
    if ($script:StepResults.Count -gt 0) {
        $blocked = $script:StepResults | Where-Object {
            $_.Status -eq "ERROR" -or
            ($_.Status -eq "WARN" -and $_.Step -ne "success.txt")
        }
        $mastResult = if (-not $blocked) { 'OK' } else { 'Error' }
        Set-MastStatus -Name $target -Status $mastResult
    }
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

$btnWinRS.Add_Click({ Show-WinRSSession })


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

$btnSccmRepair.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "SCCM Repair / Reinstall en '$target'..." `
        -Action    { Invoke-SccmRepair -ComputerName $target }
})

$btnGpUpdate.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "GPUpdate /force en '$target'..." `
        -Action {
            Write-Sep
            Write-Info "gpupdate /force en '$target'..."
            $res = Invoke-RemoteGpupdate -ComputerName $target
            Write-Sep; Append-Output "" $script:White
            if ($res -and $res.Status -eq "OK") {
                Write-Ok "gpupdate completado ($($res.Details))."
                Set-Status "GPUpdate OK en '$target'" ([System.Drawing.Color]::LightGreen)
            } else {
                $detail = if ($res) { $res.Details } else { "Sin respuesta remota" }
                Write-Warn "gpupdate: $detail."
                Set-Status "GPUpdate WARN en '$target'" ([System.Drawing.Color]::Yellow)
            }
        }
})

$btnSccmCycles.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "Ciclos SCCM en '$target'..." `
        -Action {
            Write-Sep
            Write-Info "Ciclos de politicas SCCM en '$target'..."
            $result = Invoke-LocalOrRemote -ComputerName $target -ScriptBlock $script:SccmCyclesBlock
            Write-Sep; Append-Output "" $script:White
            if ($result) {
                if ($result.Steps) { Write-StepList $result.Steps }
                $statusTxt = switch ($result.Status) {
                    "OK"    { "Ciclos completados OK" }
                    "WARN"  { "Ciclos con avisos"     }
                    default { "Ciclos con errores"    }
                }
                $statusCol = switch ($result.Status) {
                    "OK"    { [System.Drawing.Color]::LightGreen }
                    "WARN"  { [System.Drawing.Color]::Yellow     }
                    default { [System.Drawing.Color]::Tomato     }
                }
                Set-Status "$statusTxt en '$target'" $statusCol
            } else {
                Write-Warn "Sin respuesta del cliente SCCM. Verifica que CcmExec esta activo."
                Set-Status "Sin respuesta SCCM en '$target'" ([System.Drawing.Color]::Yellow)
            }
        }
})

$btnRepair.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Set-Progress 0 ""
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "Reparacion del sistema en '$target'..." `
        -Action    { Invoke-RemoteRepair -ComputerName $target }
    Set-Status "Finalizado" ([System.Drawing.Color]::LightGreen)
})

$btnChkdsk.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "ChkDsk /r en '$target'..." `
        -Action    { Invoke-RemoteChkdsk -ComputerName $target }
    Set-Status "Finalizado" ([System.Drawing.Color]::LightGreen)
})

$btnCleanup.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "Limpieza corporate en '$target'..." `
        -Action    { Invoke-CorporateCleanup -ComputerName $target }
})

$btnPerfilazo.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "Perfilazo en '$target'..." `
        -Action    { Invoke-Perfilazo -ComputerName $target }
})

$btnPerfilRestore.Add_Click({
    $target = Get-ValidComputer; if (-not $target) { return }
    Invoke-ActionButton -ComputerName $target -UseCancel $false `
        -StatusMsg "Restaurar Perfilazo en '$target'..." `
        -Action    { Invoke-PerfilRestore -ComputerName $target }
})

$btnRobocopy.Add_Click({ Show-RobocopyForm })

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
    $exists = $script:lvEquipos.Items | Where-Object { $_.Tag -eq $input }
    if ($exists) {
        [System.Windows.Forms.MessageBox]::Show(
            "'$input' ya esta en la lista.",
            "Duplicado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $item = Add-EquipoToList $input
    Save-EquipoList
    Update-EquipoCard $item
})

$btnRemoveEquipo.Add_Click({
    if ($script:lvEquipos.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Haz clic en un equipo de la lista para seleccionarlo primero.",
            "Nada seleccionado",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    $script:lvEquipos.Items.RemoveAt($script:lvEquipos.SelectedIndices[0])
    Save-EquipoList
})

$btnRefreshEquipos.Add_Click({
    if ($script:lvEquipos.Items.Count -eq 0) { return }
    Set-Status "Refrescando estado de equipos..." ([System.Drawing.Color]::Yellow)
    Refresh-EquipoEstados
    Set-Status "Listo" $white
})

# Enter en campo equipo -> ping
$txtEquipo.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") { $btnPing.PerformClick() }
})

$form.Add_Shown({
    Append-Output "  Herramienta de Administracion Remota v2.17.1" ([System.Drawing.Color]::FromArgb(0, 190, 255))
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
