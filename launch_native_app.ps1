$env:RAS_NATIVE_MODE = "sharepoint"
$env:RAS_NATIVE_SHAREPOINT_URL = "https://rcgmail-my.sharepoint.com/:f:/r/personal/gaurav_handoo_myridius_com/Documents/FInance/Solutions%20Revenue?csf=1&web=1&e=9RMo4c"

$Python = "C:\venvs\ra-desktop\Scripts\python.exe"
$ScriptDir = $PSScriptRoot

Set-Location $ScriptDir
& $Python -m native_app.main
