' Silent background launcher for the resume optimizer.
' Starts uvicorn (IPv4 + IPv6) in HIDDEN windows, logs appended under logs\.
'
' IMPORTANT -- the interpreter path below is ABSOLUTE on purpose.
' All project dependencies (fastapi / uvicorn / pymysql / python-docx / fitz /
' pdf2docx / pywin32 ...) live ONLY in D:\python (CPython 3.9.7).
' Resolving a bare "python" through PATH can pick up a different interpreter
' (e.g. WorkBuddy's managed 3.13, which has none of those packages). Because
' this script runs the servers with a HIDDEN window, that failure is totally
' silent -- no console, no error, just a dead service. Do NOT revert this
' back to bare "python".
Option Explicit

Dim sh, fso, dirPath, py, logsDir
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

dirPath = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = dirPath

py = "D:\python\python.exe"

If Not fso.FileExists(py) Then
    MsgBox "Python interpreter not found:" & vbCrLf & py & vbCrLf & vbCrLf & _
           "Project dependencies are installed only in this 3.9.7 interpreter," & vbCrLf & _
           "so the service cannot be started." & vbCrLf & vbCrLf & _
           "Fix the path in start_background.vbs, then run again.", _
           vbCritical, "Resume Optimizer - start failed"
    WScript.Quit 1
End If

logsDir = dirPath & "\logs"
If Not fso.FolderExists(logsDir) Then fso.CreateFolder(logsDir)

' 单进程启动：--host 0.0.0.0 已覆盖 IPv4 与 IPv6，无需再起第二个进程。
' 此前同时监听 0.0.0.0:8000 与 [::1]:8000 会起两个独立 uvicorn 进程，
' 而 token/session 只存在进程内存里 —— IPv4 侧登录、IPv6 侧刷新就会 401。
sh.Run "cmd /c """ & py & """ -m uvicorn main:app --host 0.0.0.0 --port 8000 >> logs\server.log 2>&1", 0, False
