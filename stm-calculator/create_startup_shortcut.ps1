$wshell = New-Object -ComObject WScript.Shell
$startupFolder = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\Startup')
$shortcutPath = [System.IO.Path]::Combine($startupFolder, 'STM_Calculator.lnk')
$shortcut = $wshell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = 'd:\Work Anti\stm-calculator\start_server_hidden.vbs'
$shortcut.WorkingDirectory = 'd:\Work Anti\stm-calculator'
$shortcut.Save()
Write-Host "Startup shortcut created at: $shortcutPath"
