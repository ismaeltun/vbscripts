Option Explicit

Include "Util.vbs"

Const Logger_LOG_LEVEL_DEBUG = 1
Const Logger_LOG_LEVEL_INFO = 2
Const Logger_LOG_LEVEL_WARN = 3
Const Logger_LOG_LEVEL_ERROR = 4
Const Logger_LOG_LEVEL_FATAL = 5

Const Logger_LOG_TYPE_FILE = 1
Const Logger_LOG_TYPE_EVENTLOG = 2

Class Logger
	private intLogLevel
	private intLogType
	private objWshShell
	Private objFso
	private strLogFilePath
	private objUtil

	private Sub Class_Initialize
		intLogLevel = Logger_LOG_LEVEL_INFO
		intLogType = Logger_LOG_TYPE_FILE

		Set objWshShell = WScript.CreateObject("WScript.Shell")
		Set objFso = CreateObject("Scripting.FileSystemObject")
		Set objUtil = New Util
	End Sub

	private Sub Class_Terminate
		On Error Resume Next
		Set objWshShell = Nothing
		Set objFso = Nothing
		Set objUtil = Nothing
	End Sub

	Public Sub SetLogLevel(ByVal strLogLevel)
		If strLogLevel = "DEBUG" Then
			intLogLevel = Logger_LOG_LEVEL_DEBUG
		ElseIf strLogLevel = "INFO" Then
			intLogLevel = Logger_LOG_LEVEL_INFO
		ElseIf strLogLevel = "WARN" Then
			intLogLevel = Logger_LOG_LEVEL_WARN
		ElseIf strLogLevel = "ERROR" Then
			intLogLevel = Logger_LOG_LEVEL_ERROR
		ElseIf strLogLevel = "FATAL" Then
			intLogLevel = Logger_LOG_LEVEL_FATAL
		Else
			Err.Raise(5)
		End If
	End Sub

	Public Sub SetLogType(ByVal strLogType)
		If strLogType = "FILE" Then
			intLogType = Logger_LOG_TYPE_FILE
		ElseIf strLogType = "EVENTLOG" Then
			intLogType = Logger_LOG_TYPE_EVENTLOG
		Else
			Err.Raise(5)
		End If
	End Sub

	Public Sub SetLogFilePath(ByVal logFilePath)
		strLogFilePath = logFilePath
	End Sub

	Public Function IsDebugEnabled
		If intLogLevel <= Logger_LOG_LEVEL_DEBUG Then
			IsDebugEnabled = True
		Else
			IsDebugEnabled = False
		End If
	End Function

	Public Function IsInfoEnabled
		If intLogLevel <= Logger_LOG_LEVEL_INFO Then
			IsInfoEnabled = True
		Else
			IsInfoEnabled = False
		End If
	End Function

	Public Function IsWarnEnabled
		If intLogLevel <= Logger_LOG_LEVEL_WARN Then
			IsWarnEnabled = True
		Else
			IsWarnEnabled = False
		End If
	End Function

	Public Function IsErrorEnabled
		If intLogLevel <= Logger_LOG_LEVEL_ERROR Then
			IsErrorEnabled = True
		Else
			IsErrorEnabled = False
		End If
	End Function

	Public Function IsFatalEnabled
		If intLogLevel <= Logger_LOG_LEVEL_FATAL Then
			IsFatalEnabled = True
		Else
			IsFatalEnabled = False
		End If
	End Function

	Public Sub Debug(ByVal strMessage)
		On Error Resume Next
		If IsDebugEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_DEBUG, strMessage)
		End If
	End Sub

	Public Sub Info(ByVal strMessage)
		On Error Resume Next
		If IsInfoEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_INFO, strMessage)
		End If
	End Sub

	Public Sub Warn(ByVal strMessage)
		On Error Resume Next
		If IsWarnEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_WARN, strMessage)
		End If
	End Sub

	Public Sub Error(ByVal strMessage)
		On Error Resume Next
		If IsErrorEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_ERROR, strMessage)
		End If
	End Sub

	Public Sub Fatal(ByVal strMessage)
		On Error Resume Next
		If IsFatalEnabled Then
			Call WriteLog(Logger_LOG_LEVEL_FATAL, strMessage)
		End If
	End Sub

	Private Sub WriteLog(ByVal logLevel, ByVal strMessage)
		On Error Resume Next
		If intLogType = Logger_LOG_TYPE_FILE Then
			Call WriteLogFile(logLevel, strMessage)
		ElseIf intLogType = Logger_LOG_TYPE_EVENTLOG Then
			Call WriteEventLog(logLevel, strMessage)
		End If
	End Sub

	Private Sub WriteEventLog(ByVal logLevel, ByVal strMessage)
		On Error Resume Next

		Dim intEventLogLevel

		If logLevel = Logger_LOG_LEVEL_DEBUG Then
			intEventLogLevel = 4
		ElseIf logLevel = Logger_LOG_LEVEL_INFO Then
			intEventLogLevel = 4
		ElseIf logLevel = Logger_LOG_LEVEL_WARN Then
			intEventLogLevel = 2
		ElseIf logLevel = Logger_LOG_LEVEL_ERROR Then
			intEventLogLevel = 1
		ElseIf logLevel = Logger_LOG_LEVEL_FATAL Then
			intEventLogLevel = 1
		End If

		objWshShell.LogEvent intEventLogLevel, strMessage
	End Sub

	Private Sub WriteLogFile(ByVal logLevel, ByVal strMessage)
		On Error Resume Next

		Const ForReading = 1, ForAppending = 8
	    Dim dateNow
	    dateNow = Now
		
		Dim objLogFile
		Set objLogFile = objFso.OpenTextFile(strLogFilePath, ForAppending, True)
		Dim strLogLevel
		If logLevel = Logger_LOG_LEVEL_DEBUG Then
			strLogLevel = "DEBUG"
		ElseIf logLevel = Logger_LOG_LEVEL_INFO Then
			strLogLevel = "INFO"
		ElseIf logLevel = Logger_LOG_LEVEL_WARN Then
			strLogLevel = "WARN"
		ElseIf logLevel = Logger_LOG_LEVEL_ERROR Then
			strLogLevel = "ERROR"
		ElseIf logLevel = Logger_LOG_LEVEL_FATAL Then
			strLogLevel = "FATAL"
		End If

		objLogFile.WriteLine(objUtil.FormatDate(dateNow, "YYYY/MM/DD") & Space(1) & objUtil.FormatDate(dateNow, "HH24:MI:SS") & " [" & strLogLevel & "] " & strMessage)
	    objLogFile.Close
		Set objLogFile = Nothing
	End Sub
End Class

Sub Include(ByVal FilePath)
	On Error Resume Next

	Dim objFso, objTextStream

    Set objFso = WScript.CreateObject("Scripting.FileSystemObject")

	Set objTextStream = objFso.OpenTextFile(FilePath, "1", False)

	ExecuteGlobal(objTextStream.ReadAll)

    objTextStream.Close
	Set objTextStream = Nothing
	Set objFso = Nothing
End Sub
