Dim WshShell, oExec
Set WshShell = CreateObject("WScript.Shell")

' 이미 실행 중이면 종료
Dim oWMI, colProc
Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
Set colProc = oWMI.ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%uvicorn%main:app%'")
If colProc.Count > 0 Then
    WScript.Quit
End If

' 서버 시작 (창 없이)
WshShell.CurrentDirectory = "D:\Work Anti\stm-calculator"
WshShell.Run """C:\Users\220216\AppData\Local\Programs\Python\Python314\python.exe"" -m uvicorn main:app --host 0.0.0.0 --port 8502", 0, False
