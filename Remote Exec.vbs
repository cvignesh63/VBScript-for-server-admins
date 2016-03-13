Option Explicit
Dim objWMIService, objProcess
Dim strShell, objProgram, strComputer, strExe, strInput
strExe = "calc.exe"

Do
strComputer = (InputBox(" ComputerName to Run Script", "Computer Name"))
If strComputer <> "" Then
strInput = True
End if
Loop until strInput = True

set objWMIService = getobject("winmgmts://"& strComputer & "/root/cimv2")
Set objProcess = objWMIService.Get("Win32_Process")
Set objProgram = objProcess.Methods_("Create").InParameters.SpawnInstance_
objProgram.CommandLine = strExe

Set strShell = objWMIService.ExecMethod( "Win32_Process", "Create", objProgram)

WScript.echo "Executed: " & strExe & " on " & strComputer
WSCript.Quit
