Set objShell = CreateObject("WScript.Shell")
Wscript.Echo objShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\" &_
                   "JavaSoft\Java Runtime Environment\1.8\JavaHome")

'check file version of java.exe
javaHome = objShell.Environment.item("JAVA_HOME")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Wscript.Echo objFSO.GetFileVersion(javaHome & "\bin\java.exe")
