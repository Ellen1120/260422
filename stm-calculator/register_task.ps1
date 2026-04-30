$action = New-ScheduledTaskAction `
    -Execute "wscript.exe" `
    -Argument "`"d:\Work Anti\stm-calculator\start_server_hidden.vbs`"" `
    -WorkingDirectory "d:\Work Anti\stm-calculator"

$trigger = New-ScheduledTaskTrigger -AtStartup

$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit 0 -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest

Register-ScheduledTask `
    -TaskName "STM Calculator Server" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Force

Write-Host "완료! STM Calculator Server 작업이 등록되었습니다." -ForegroundColor Green
