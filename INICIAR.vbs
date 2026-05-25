Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
WshShell.Run "powershell.exe -ExecutionPolicy Bypass -NoExit -Command ""Set-Location '" & Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1) & "'; python -m pip install flask openpyxl -q; Start-Job { Start-Sleep 3; Start-Process 'http://localhost:5000' } | Out-Null; python app.py""", 1, False
