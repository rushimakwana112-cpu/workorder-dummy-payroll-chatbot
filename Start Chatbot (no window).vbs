' Starts the Payroll Chatbot WITHOUT a black CMD window:
'   Blue Panther site  -> http://<this-PC>:8080   (web_server.py, restarted automatically if it stops)
'   (the old Streamlit site on :8501 is no longer started; first admin ID: Reset Password.bat)
' Double-click this file (it also runs at Windows sign-in). To stop it, double-click Stop Chatbot.bat.
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dir
flag = dir & "\stop.flag"
If fso.FileExists(flag) Then fso.DeleteFile flag

' keep the new site running: start it, wait until it stops, start it again (unless Stop Chatbot.bat asked to stop)
Do
    sh.Run "cmd /c python web_server.py >> web_server.log 2>&1", 0, True
    If fso.FileExists(flag) Then Exit Do
    WScript.Sleep 5000
    If fso.FileExists(flag) Then Exit Do
Loop
If fso.FileExists(flag) Then fso.DeleteFile flag
