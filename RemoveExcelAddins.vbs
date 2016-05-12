Option Explicit

Const TestAddIn = "C:\atest.xla"

Sub CreateAddIn()
  Dim W As Workbook
  Set W = Workbooks.Add
  W.IsAddin = True
  W.SaveAs TestAddIn, xlAddIn
  W.Close
End Sub

Sub LoadAddIn()
  AddIns.Add TestAddIn, False
  AddIns("atest").Installed = True
End Sub

Sub RemoveAddIn()
  Dim fso As FileSystemObject
  Dim ts As TextStream
  Dim Script As String, ScriptFile As String
  Dim A As AddIn
  Dim objShell
 
  'Access the addin to remove
  Set A = AddIns("atest")
  A.Installed = False

Script = _
"On Error Resume Next" & vbCrLf & _
"'Wait until Excel is closed" & vbCrLf & _
"WScript.Sleep 1000" & vbCrLf & _
"'Here is the place where Excel stores the AddIns list" & vbCrLf & _
"RegPath = ""Software\Microsoft\Office\[Version]\Excel\Add-in Manager""" & vbCrLf & _
"'Delete the addin from the list" & vbCrLf & _
"Set oReg = GetObject(""winmgmts:{impersonationLevel=impersonate}!"" & _" & vbCrLf & _
"""\\.\root\default:StdRegProv"")" & vbCrLf & _
"oReg.DeleteValue &H80000001, RegPath, ""[AddIn]""" & vbCrLf & _
"'Restart Excel" & vbCrLf & _
"Set objShell = CreateObject(""Wscript.Shell"")" & vbCrLf & _
"objShell.Run ""excel.exe""" & vbCrLf & _
"'Delete this script" & vbCrLf & _
"Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
"fso.DeleteFile Wscript.ScriptFullName"

  Script = Replace(Script, "[Version]", Application.Version)
  Script = Replace(Script, "[AddIn]", A.FullName)

  'Create the VBScript
  Set fso = CreateObject("Scripting.FileSystemObject")
  ScriptFile = Replace(fso.GetTempName, ".tmp", ".vbs")
  ScriptFile = fso.BuildPath(Environ$("temp"), ScriptFile)
  Set ts = fso.CreateTextFile(ScriptFile)
  ts.Write Script
  ts.Close
 
  'Run it and close Excel
  Set objShell = CreateObject("Wscript.Shell")
  objShell.Run ScriptFile
  Application.DisplayAlerts = False
  Application.Quit
End Sub