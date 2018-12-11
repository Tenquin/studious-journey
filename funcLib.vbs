'=============== FUNCTION LIBRARY ========================
'Check to see if there is a USB CD drive connected. Returns true or fale boolean.
Function HasUSBCD()
	return = False
	On Error Resume Next
	strQuery = "SELECT * FROM Win32_CDROMDrive WHERE DeviceID LIKE 'USBSTOR%'"
	Set objWMIService = GetObject("winmgmts:\\.\ROOT\CIMV2")
	Set colItems = objWMIService.ExecQuery(strQuery)
	count = 0
	For Each objItem in colItems
		return = True 
	Next
	HasUSBCD = return
End Function

'Gets MFG of Motherboard
Function MBMFG() 
	Set WMI = GetObject("WinMgmts:")
	Set objs = WMI.InstancesOf("Win32_BaseBoard")
	For Each obj In objs
 		 sAns = sAns & obj.Manufacturer
 	If sAns < objs.Count Then sAns = sAns & ","
	Next
	MBMFG = Left(sAns,16)
End Function

'Run an app
Sub RunApp(appName,waitState)
    Set objShell = CreateObject("Wscript.Shell")
	objShell.Run appName,2,waitState
End Sub
 
 'Copy a file
Function CopyFile(strFile, strDest)
	CopyFile = "INIT"
	On Error Resume Next
	set objFSO=CreateObject("Scripting.FileSystemObject")
	'set last parameter to TRUE if you want to overwrite an existing file with the same name
	objFSO.CopyFile strFile,strDest,TRUE
	if err.number=0 Then
		CopyFile = "SUCCESS"
	Else
		CopyFile = "FAIL:Err:" & err.number & " " & err.Decsription & " :srcPath:" & strFile & " :destPath:" & strDest
	end If
End Function

'reads a file's contents into memory
Function ReadFile(filePath)
	return = ""
	On Error Resume Next
	Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
	return = oFileSystemObject.OpenTextFile(filePath,1).ReadAll
	ReadFile = return
End Function

'Create a file with specic content
Function CreateAFile(strFileName,strContent)
	Const AppendIfExists = True
	Const ForReading = 1
	Const ForAppending = 8
	Const ForWriting = 2
	'on error resume next
	Set FS=CreateObject("Scripting.FileSystemObject")
	Set fFile = FS.OpenTextFile(strFileName,ForWriting,AppendIfExists)
	fFile.Write(strContent)
	fFile.Close
End Function

'HTA version: Gets current operating directory based on current file name
Function GetCurDir()
	strPattern = "(\w:\\[\w*\s–\-\&\\\.]*)"
	strFileFind = "([\w-_\s]*\.hta$)"
	Set RegEx = New RegExp
	RegEx.IgnoreCase = True
	RegEx.Global=False
	RegEx.Pattern=strPattern
	Set strFound = RegEx.Execute(document.url)
		For Each key in strFound
			return1 = key.value
		Next
	RegEx.Pattern=strFileFind
	Set strFound = RegEx.Execute(return1)
		For Each key in strFound
			return2 = key.value
		Next 
	GetCurDir = Replace(return1,return2,"")
End Function

'Deletes a file or files
Function DeleteFolder(strFile)
	On Error Resume Next
	dim objFSO
	set objFSO=CreateObject("Scripting.FileSystemObject")
	'to force deletion, set to TRUE otherwise set to FALSE
	objFSO.DeleteFolder strFile,TRUE
End Function

'reboot system
Function sysShutdown()
	Set OpSysSet = GetObject("winmgmts:{authenticationlevel=Pkt," _
		 & "(Shutdown)}").ExecQuery("select * from Win32_OperatingSystem where "_
		 & "Primary=true")
	for each OpSys in OpSysSet
		retVal = OpSys.Reboot()
	next
End Function

'Parse the logfile text
Function ParseLog(strData)	on Error Resume Next
	SplitLine = Split(strData,"$")
	RowLen = Ubound(SplitLine)
		'Wscript.Echo RowLen
		For ii = 0 to RowLen
			SplitData = Split(SplitLine(ii),",")				
				return = return & SplitData(1) & vbLF 	
		Next
	ParseLog = return
End Function

'Returns boolean of whether a file exists
Function FileExists(strFile)
	On Error Resume Next
	return = False
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFile) Then
		return = True
	Else
		return = False
	End If
	FileExists=return
End Function

'Open a system control panel
Function OpenCP(strName)
	set objShell = CreateObject("shell.application")
	objshell.ControlPanelItem(strName)
	set objShell = Nothing
End Function

'Send a logFile message
Function SendLog(strLogfile,strMessage,strEvent)
	Const AppendIfExists = True
	Const ForReading = 1
	Const ForAppending = 8
	on error resume next
	Set FS=CreateObject("Scripting.FileSystemObject")
		Set fFile = FS.OpenTextFile(strLogfile,ForAppending,AppendIfExists)
	strContent = Date() & "." & Time() & "," & strMessage &	"," & strEvent & "$" 
		fFile.Write(strContent & vbCrlf)
	fFile.Close
		SendLog = strContent
End Function

'Deletes a file or files
Function DeleteFile(strFile)
	On Error Resume Next
	dim objFSO
	set objFSO=CreateObject("Scripting.FileSystemObject")
	'to force deletion, set to TRUE otherwise set to FALSE
	objFSO.DeleteFile strFile,TRUE
	if err.number=0 Then
		return = "SUCCESS"
	Else
		return = "FAIL:Err:" & err.number & " " & err.Description & " File:" & strFile
	end If
	DeleteFile = return
End Function

'expands an enviromental paths for windows Enviromental Variables
Function EnvPath(strDosVariable)
	return = ""
	Set wshShell = CreateObject( "WScript.Shell" )
	return = wshShell.ExpandEnvironmentStrings(strDosVariable)
	EnvPath = return
End Function