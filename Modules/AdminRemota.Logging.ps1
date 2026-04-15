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
