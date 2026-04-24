# Script-toolKit — Contexto para Claude Code

## Control de versiones — OBLIGATORIO

**Nunca commitear directamente en `main`.**

Flujo correcto para cada tarea:
1. Crear o cambiar a la rama de feature correspondiente
2. Desarrollar y commitear en esa rama
3. Hacer `git push -u origin <rama>`
4. Hacer merge a `main` solo cuando el usuario lo pida explicitamente

### Rama activa por defecto
```
claude/refactor-powershell-winforms-KTvwc
```
Si no existe localmente: `git checkout -B claude/refactor-powershell-winforms-KTvwc`

### Antes de cualquier commit, verificar:
```
git branch --show-current   # debe ser la rama feature, NO main
```

---

## Estructura del proyecto

Herramienta WinForms de administracion remota en PowerShell 5.1.

```
AdminRemota.ps1              # Punto de entrada: GUI, botones, event handlers
Modules/
  AdminRemota.Gui.ps1        # New-FlatButton, New-GroupPanel, Show-*Form, Invoke-ActionButton
  AdminRemota.Logging.ps1    # Append-Output, Write-Info/Ok/Warn/Fail, Write-Sep, Set-Status
  AdminRemota.Remote.ps1     # Invoke-LocalOrRemote, Test-IsLocal, SccmCyclesBlock, Get-TargetNetworkZone
  AdminRemota.Steps.ps1      # Invoke-Step, Reset/Add/Show-StepResults
  AdminRemota.Master.ps1     # Invoke-MasterCheck
  AdminRemota.Sccm.ps1       # Invoke-SoftwareCheck
  AdminRemota.System.ps1     # Invoke-SystemInfo, Invoke-UsbDriverClean, reparacion, chkdsk...
  AdminRemota.Perfilazo.ps1  # Invoke-Perfilazo, Invoke-PerfilRestore
masterParaRevision.txt       # Script monolitico pre-modularizacion (referencia historica)
```

## Patrones clave

- Botones nuevos: `New-FlatButton` → añadir a `$script:ActionButtons` si deben bloquearse durante tareas
- Ventanas secundarias: `Show-*Form` en `AdminRemota.Gui.ps1`, usar `.GetNewClosure()` en event handlers
- Ejecucion remota: siempre via `Invoke-LocalOrRemote`, nunca `Invoke-Command` directo
- Errores WinRM: no terminantes por defecto (`$PSDefaultParameterValues['Invoke-Command:ErrorAction']='SilentlyContinue'`)
- Pasos de masterizacion: envolver en `Invoke-Step` para que los errores queden en el resumen, no en la GUI
