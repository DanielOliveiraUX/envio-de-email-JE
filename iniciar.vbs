Dim oShell, sDir
Set oShell = CreateObject("WScript.Shell")
sDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\")) & "Projeto"
oShell.CurrentDirectory = sDir
oShell.Run "cmd /c cd /d """ & sDir & """ && pip install -r requirements.txt -q && python app.py", 0, False
Set oShell = Nothing
