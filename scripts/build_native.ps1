param(
    [string]$Name = "RevenueAnalysisDesktop",
    [string]$PythonExe = ""
)

$ErrorActionPreference = "Stop"

if (-not $PythonExe) {
    $LocalVenvPython = Join-Path $PSScriptRoot "..\.venv\Scripts\python.exe"
    if (Test-Path $LocalVenvPython) {
        $PythonExe = $LocalVenvPython
    }
    elseif (Test-Path "C:\venvs\ra-desktop\Scripts\python.exe") {
        $PythonExe = "C:\venvs\ra-desktop\Scripts\python.exe"
    }
    else {
        throw "Python executable not found. Pass -PythonExe explicitly."
    }
}

Set-Location (Join-Path $PSScriptRoot "..")
& $PythonExe -m PyInstaller --noconfirm --windowed --name $Name native_app/main.py
