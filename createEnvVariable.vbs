  Set WshShell = WScript.CreateObject("WScript.Shell")
  If WScript.Arguments.length = 0 Then
  Set ObjShell = CreateObject("Shell.Application")
  ObjShell.ShellExecute "wscript.exe", """" & _
  WScript.ScriptFullName & """" &_
  " RunAsAdministrator", , "runas", 1
  Wscript.Quit
  End if

If WScript.Arguments.Count = 0 Then
    WScript.Echo "Usage: cscript createEnvVariable JAVA_HOME ""PathToJavaHome"""
    WScript.Echo "Missing parameters expected VAR [the environment var to set] and VAR [the parameter to set it to]"

Else
	var = WScript.Arguments(0)
	val = WScript.Arguments(1)
	
'	Set objVarClass = GetObject( "winmgmts://./root/cimv2:Win32_Environment" )
'	Set objVar      = objVarClass.SpawnInstance_
'	objVar.Name          = var
'	objVar.VariableValue = val
'	objVar.UserName      = "<SYSTEM>"
'	objVar.Put_
'	WScript.Echo "Created environment variable " & strVarName
'	Set objVar      = Nothing
'	Set objVarClass = Nothing
	
	Set wshShell = CreateObject( "WScript.Shell" )
	Set wshSystemEnv = wshShell.Environment( "SYSTEM" )
	WScript.Echo "Setting env var: " & var & " to " & val
	wshSystemEnv( var ) = val
	Set wshSystemEnv = Nothing
	Set wshShell     = Nothing
End If
