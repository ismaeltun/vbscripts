Dim group
Dim publicDir
Set Args = WScript.Arguments
Const NO_RIGHTS_TOKEN = ")n"
Const PATH_TOKEN = ":\"
localizedUserGroupName = ""

group = Args(0)

'Get the localized user group name.
Set objWMIService = GetObject( _
    "winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_Group Where LocalAccount = True And SID = 'S-1-5-32-545'")   
if colItems.Count = 1 Then
 localizedUserGroupName = colItems.ItemIndex(0).Name
End if

' Check to see if the argument is a path or a user group
position = InStr(1, group, PATH_TOKEN, 1)

' It's a path, see if the path is in the public folder
if ( position > 0 ) Then 

  Set Shell = CreateObject("WScript.Shell")
  Set Environment = Shell.Environment("Process")
  publicDir = Environment("Public")

  position = InStr(1, group, publicDir, 1)

  ' If position < 1, then the path is not in the public directory
  if position < 1 Then
    Wscript.Quit 0

  ' This path is in the public directory, the users group will have rights to it
  else
    Wscript.Quit 1
  End if

' It's not a path, it's a user group
else

  ' In theory we should always find the "users" group name. If we don't then a user definitely
  ' has no access.
  if localizedUserGroupName = "" Then
    Wscript.Quit 0
  End if
  
  ' Check to see if the group contains permissions for the specified user group
  position = InStr(1, group, localizedUserGroupName, 1)

  ' If thegroup is not listed, then the group has no rights
  if position < 1 Then WScript.Quit 0

  ' If the group is listed, make sure their rights aren't set to "no rights" 
  position = InStr(1, group, NO_RIGHTS_TOKEN, 1)

  ' If position < 1, the string that specifies the group has no rights isn't in the string, so the group must have some rights
  if position < 1 Then 
    Wscript.Quit 1

  ' User has no rights to this directory
  else
    Wscript.Quit 0
  End if  
End if


