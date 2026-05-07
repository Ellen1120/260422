Set WshShell = CreateObject("WScript.Shell")
strPath = "d:\Work Anti\stm-calculator"
WshShell.CurrentDirectory = strPath
WshShell.Run "C:\Users\220216\AppData\Local\Programs\Python\Python314\python.exe -m uvicorn main:app --host 0.0.0.0 --port 8502", 0, False
