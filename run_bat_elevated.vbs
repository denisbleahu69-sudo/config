Set objShell = CreateObject("Shell.Application")
Set objWshShell = CreateObject("WScript.Shell")

' Get arguments: batPath, workingDir, batFileName
Dim args
Set args = WScript.Arguments

If args.Count < 3 Then
    WScript.Echo "Usage: run_bat_elevated.vbs <batPath> <workingDir> <batFileName>"
    WScript.Quit
End If

Dim batPath, workingDir, batFileName
batPath = args(0)
workingDir = args(1)
batFileName = args(2)

' Change to the working directory
objWshShell.CurrentDirectory = workingDir

' Run the batch file as administrator (1 = show window, 0 = hide)
objShell.ShellExecute "cmd.exe", "/k """ & batPath & """", workingDir, "runas", 1

' Wait for cmd.exe processes to finish (check every second)
Do
    WScript.Sleep 1000
    Set colProcesses = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT * FROM Win32_Process WHERE Name='cmd.exe'")
    Dim found
    found = False
    For Each proc In colProcesses
        If InStr(proc.CommandLine, batFileName) > 0 Then
            found = True
            Exit For
        End If
    Next
    If Not found Then Exit Do
Loop

