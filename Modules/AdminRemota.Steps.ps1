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
        # Invoke-LocalOrRemote usa ErrorAction=Stop: los fallos de red/WinRM lanzan
        # excepcion y caen al catch de abajo. El path null==WARN es seguridad adicional
        # para scriptblocks que no devuelven valor (return sin argumento).
        $status  = if     ($res -is [hashtable] -and $res.Status) { $res.Status }
                   elseif ($null -eq $res)                         { "WARN"      }
                   else                                            { "OK"        }
        $details = if ($res -is [hashtable] -and $res.Details) { ($res.Details | Out-String).Trim() }
                   elseif ($null -eq $res)                      { "Sin respuesta (scriptblock no devolvio valor)" }
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
