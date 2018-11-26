'=============== FUNCTION LIBRARY ========================
'Execute Code from another VBS file ' performs error validation on execute. 
Function IncludeCode(strFile)
	On Error Resume Next
	return = 99
	Const ForReading = 1
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then
		return = "IncludeCode Error: File not found: " & strFile 
	Else
		Set objFile = objFSO.OpenTextFile(strFile, ForReading)
		If Err <> 0 Then
			return = "IncludeCode Unable to open input text file " & strFile & vbCrLf & _
			"Error: " & Err.Number & vbCrLf & Err.Description
			Err.Clear
		else
			ExecuteGlobal objFile.ReadAll()
			If Err <> 0 Then 
				return = "IncludeCode Execute Failure in [" & strFile & "]" & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description & "Source:" & vbCrLf & Err.Source
				Err.Clear
			else
				return = 0
			end if		
		End if		
	End If
	IncludeCode = return
End Function

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

'Attempts to pull Serial number from system BIOS
Function GetSerial() 
	Set WMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set objs = WMI.InstancesOf("Win32_Bios")
	For Each obj In objs
  		sAns = sAns & obj.SerialNumber
 		If sAns < objs.Count Then sAns = sAns & ","
	Next
		GetSerial = Left(Replace(sAns," ",""),20)
	End Function

'Gets MFG of Motherboard
Function GetMFG() 
	Set WMI = GetObject("winmgmts:\\.\ROOT\CIMV2")
	Set objs = WMI.InstancesOf("Win32_BaseBoard")
	For Each obj In objs
 		 sAns = sAns & obj.Manufacturer
 	If sAns < objs.Count Then sAns = sAns & ","
	Next
	GetMFG = Left(sAns,16)
End Function

'Gets current model name of system
Function GetModel()
	return= ""
	On Error Resume Next
	Const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20
	Set objWMIService = GetObject("winmgmts:\\.\ROOT\CIMV2")
	strQuery = "SELECT * FROM Win32_ComputerSystem"
	Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
		For Each objItem in colItems
		   return = objItem.Model
		Next
	GetModel = RTrim(LTrim(Replace(return," ","")))
End Function

'Run an app, waitstate Boolean True False waits for app to termiante
Function RunApp(appName,waitState)
	'Wscript.echo appName
	return = "INIT"
    Set objShell = CreateObject("Wscript.Shell")
	exe1 = objShell.Run(appName,2,waitState)
	If exe <> 0 Then
		return = "RunApp_Fail:" & exe1 & " " & err.Description & " :App:" & appName
	else 
		return = "SUCCESS"
	end if
		RunApp = return
End function
 
'Copy a file to a destination
Function CopyFile(strFile, strDest)
	return = "INIT"
	On Error Resume Next
	set objFSO=CreateObject("Scripting.FileSystemObject")
	objFSO.CopyFile strFile,strDest,TRUE 'set last parameter to TRUE if  overwrite file 
	if err.number=0 Then
		return = "SUCCESS"
	Else
		return = "CopyFileFAIL:Err:" & err.number & " " & err.Decsription & " :srcPath:" & strFile & " :destPath:" & strDest
	end If
	CopyFile = return
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

'Returns boolean of whether a file exists
Function FolderExists(strPath)
	On Error Resume Next
	return = false
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(strPath) Then
		return = true
	Else
		return = false
	End If
	FolderExists=return
End Function

'reboot system
Function sysReboot()
	Set OpSysSet = GetObject("winmgmts:{authenticationlevel=Pkt," _
		 & "(Shutdown)}").ExecQuery("select * from Win32_OperatingSystem where "_
		 & "Primary=true")
	for each OpSys in OpSysSet
		retVal = OpSys.Reboot()
	next
End Function

'Parse the logfile text
Function ParseLog(strData)	
	on Error Resume Next
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
	return = False
	On Error Resume Next
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFile) Then
		return = True
	Else
		return = False
	End If
	FileExists=return
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
		return = "DeleteFileFAIL:Err:" & err.number & " " & err.Description & " File:" & strFile
	end If
	DeleteFile = return
End Function

'expands an enviromental paths for windows Enviromental Variables
Function EnvPath(strDosVariable)
	return = "NONE"
	Set wshShell = CreateObject( "WScript.Shell" )
	return = wshShell.ExpandEnvironmentStrings(strDosVariable)
	EnvPath = return
End Function

'Stops a specific process by name
Function StopProcess(strAppname)
	return = "NOT_FOUND"
	on error resume next
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process Where Name like '%" & strAppname & "%'")
    For Each objItem In colItems        
		objItem.Terminate
		if err.Number <> 0 then
			return = "CloseAppFail:" & err.Number & " " & err.Description & "AppName:" & strAppname
		else
			return = "SUCCESS"
		End If
    Next
	StopProcess = return
End Function

'This function checks to see if a particular process CMD is running. Returns Boolean True/False. 
Function IsCMDRunning(strName)
	Dim objWMIService, objProcess, colProcess
	Dim  strList
	return = False
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colProcess = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & strName & "%'")
	For Each objProcess in colProcess
		If Len(objProcess.Name) > 0 then
			return = True
		End if
	Next
	IsCMDRunning = return
End Function

'Get string results from a command shell std out
Function GetBINResult(binFile)
	Set GetBINResult = Wscript.StdIn
	on error resume next
	Dim oShellOutputFileToRead, iErr
	Set oShellObject = CreateObject("Wscript.Shell")
	Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
	sShellRndTmpFile = "c:\temp\cmd.temp"
	oShellObject.Run "%comspec% /c " & binFile & " > " & sShellRndTmpFile, 0, True
	iErr = Err.Number
   	On Error GoTo 0
	If iErr <> 0 Then 
		GetBINResult = Err.Description
		Exit Function 
	End If 
	Wscript.Sleep 2500
	GetBINResult = oFileSystemObject.OpenTextFile(sShellRndTmpFile,1).ReadAll
	oFileSystemObject.DeleteFile sShellRndTmpFile, True 
End Function

'dismounts and ejects a drive given a drive letter path sich as "d:"
Function EjectCD(strDrivePath)
	on Error resume next
	For Each d in CreateObject("Scripting.FileSystemObject").Drives
		If d.DriveType = 4 Then
			CreateObject("Shell.Application").Namespace(17).ParseName(strDrivePath).InvokeVerb("Eject")
		End If
	Next
	d = Nothing
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set	objDrives = objWMI.ExecQuery("SELECT * FROM Win32_Volume Where DriveType = 5 and driveletter = '" & strDrivePath & "'")
	for each objDrive in objDrives
		retVal = objDrive.Dismount(true,false)
	next	
End Function

'Dismounts a volume based on drive letter such as "d:"
Function UnmountDrive(strDrivePath)
	retVal = 99
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set	objDrives = objWMI.ExecQuery("SELECT * FROM Win32_Volume Where DriveLetter = '" & strDrivePath & "'")
	for each objDrive in objDrives
		retVal = objDrive.Dismount(true,false)
	next	
	UnmountDrive = retVal
End Function

'Returns how many image files have been copies so far by reading a robocopy log file.
Function GetRoboFilesCount(strData)
	Set objRE = New RegExp
	objRE.Global     = True
	objRE.IgnoreCase = False
	objRE.Pattern    = "([\w-_]*Par[\w-_]*\.img)"
	Set colMatches = objRE.Execute(strData)
		For Each objMatch In colMatches
			return =  colMatches.Count
		Next
	GetRoboFilesCount = return	
End Function

'returns the RoboCopy Percent from log data
Function GetRoboPercent(strData)
	return = "0%"
	On Error Resume Next
	Set objRE = New RegExp
	objRE.Global     = True
	objRE.IgnoreCase = False
	objRE.Pattern    = "(\d+\.*\d\%)"
	Set colMatches = objRE.Execute(strData)
		For Each objMatch In colMatches
			return =   objMatch.Value
		Next
	GetRoboPercent = return
End Function

'gets last file copied from RoboCopy Log file
Function GetRoboLastFile(strData)
	Set objRE = New RegExp
	objRE.Global     = True
	objRE.IgnoreCase = False
	objRE.Pattern    = "([\w-_]*Par[\w-_]*\.img)"
	Set colMatches = objRE.Execute(strData)
		For Each objMatch In colMatches
			return =   objMatch.Value
			return = Replace(return," ","")
		Next
	GetRoboLastFile = return	
End Function

'RETRY LIMIT EXCEEDED from Log file
Function GetRoboError(strData)
	Set objRE = New RegExp
	objRE.Global     = True
	objRE.IgnoreCase = False
	objRE.Pattern    = "(RETRY LIMIT EXCEEDED)"
	Set colMatches = objRE.Execute(strData)
		For Each objMatch In colMatches
			return =   objMatch.Value
		Next
	if return = "" Then
		return = False
	else
		return = True
	End if
	GetRoboError = return	
End Function

'Returns if Robocopy log task has finished but is still running as well. 		
Function GetRoboState(strData)
	Set objRE = New RegExp
	objRE.Global     = True
	objRE.IgnoreCase = False
	objRE.Pattern    = "(Files\s:\s*[0-9]+)"
	return = objRE.Test(strData)
	GetRoboState = return
End Function

'Returns the root path of an existing file
Function FindPath(strPath)
	return = ""
	On Error Resume Next
	Set  objFSO = CreateObject("Scripting.FileSystemObject")
	LetArray = Array("a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z")
	idxLast = UBound( LetArray )
	For i = 0 To idxLast
		sPath = LetArray(i) & ":\"
		fPath = sPath & strPath
			'Determine if file exists
			If objFSO.FileExists(fPath)  Then
				return = sPath
			End If 
	Next
	Set objFSO = Nothing 
	FindPath = return
End Function

'Gets the drive letter with an inserted disc.
Function GetCDDrive()
	return = ""
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery( "Select * from Win32_CDROMDrive Where MediaLoaded = 'True'")
	For Each objItem in colItems
		return = objItem.Drive
	Next
	GetCDDrive = return
End Function 

'Alternmetod to get CD/DVD Drive letter.
Function GetCDDrive2()
	return = "NONE"
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery( "SELECT DeviceID FROM Win32_LogicalDisk WHERE Access = 1 AND DriveType = 5")
	For Each objItem in colItems
		return = objItem.DeviceID
  Next
  if return = "" then
    return = "NONE"
  end if
	GetCDDrive2 = return
End Function

'Function which will return an External USB HDD or USB Flash drive letter (Does not work in WinPE 10 1803)
Function GetExternalHDD()
	return = ""
	On Error Resume Next
	Set WinShell = Createobject("wscript.Shell")
	Set wmiServices = GetObject("winmgmts:{impersonationLevel=Impersonate}!//.")
	SQL = "SELECT DeviceID FROM Win32_DiskDrive WHERE PNPDeviceID like 'USBSTOR%'" 
	Set wmiDiskDrives = wmiServices.ExecQuery(SQL)
	For Each wmiDiskDrive In wmiDiskDrives
		'WScript.Echo wmiDiskDrive.Caption & " (" & wmiDiskDrive.DeviceID & ")"
		strEscapedDeviceID = Replace(wmiDiskDrive.DeviceID, "\", "\\", 1, -1, vbTextCompare)
		Set wmiDiskPartitions = wmiServices.ExecQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & strEscapedDeviceID & """} WHERE AssocClass = " & "Win32_DiskDriveToDiskPartition")
		For Each wmiDiskPartition In wmiDiskPartitions
			Set wmiLogicalDisks = wmiServices.ExecQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & wmiDiskPartition.DeviceID & """} WHERE AssocClass = " & "Win32_LogicalDiskToPartition")
	 		For Each wmiLogicalDisk In wmiLogicalDisks
				return = wmiLogicalDisk.DeviceID
			Next
		Next
	Next
	GetExternalHDD = return
End Function

'Altermate method of getting external HDD drive
Function GetExtHDD(strFolderPath)
	return = ""
	On Error Resume Next
	Const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20 'create an array of availible drives from WMI call
	strQuery = "SELECT Access,DriveType,MediaType,DeviceID,VolumeName FROM Win32_LogicalDisk where DriveType < 4 AND VolumeName <> null"
	Set objWMIService = GetObject("winmgmts:\\.\ROOT\CIMV2")
	Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
	For Each objItem in colItems
		drvArray = drvArray  & objItem.DeviceID
		drvArray = drvArray & ","    
	Next
	drvArray = Left(drvArray,Len(drvArray) -1)
	'find the foldr with the path from above array
	Set  objFSO = CreateObject("Scripting.FileSystemObject")
	LetArray = split(drvArray,",")
	idxLast = UBound(LetArray)
	For i = 0 To idxLast
		sPath = LetArray(i) 
		fPath = sPath & strFolderPath  
		'Determine if folder exists
		If objFSO.FolderExists(fPath)  Then
			return = sPath 
		End If 
	Next
	Set objFSO = Nothing 
	GetExtHDD = return
End function

'Get dism operation percent from text object
Function GetDismPercent(strData)
    return = 0
    Set objRE = New RegExp
	objRE.Global     = True
	objRE.IgnoreCase = False
	objRE.Pattern    = "(\d{1,3}\.0%)"
	Set colMatches = objRE.Execute(strData)
		For Each objMatch In colMatches
            return =  colMatches.Count
            return = objMatch.Value
		Next
        GetDismPercent = return	
End Function

'Get DISM operation state. Boolean returned True is successful. 
Function GetDismState(strData)
    Set objRE = New RegExp
	objRE.Global     = True
	objRE.IgnoreCase = False
	objRE.Pattern    = "(com\w+\ssuc\w+)"
	return = objRE.Test(strData)
	GetDismState = return	
End Function

'Shuts down WINPE enviroment
Sub Shutdown()
	on error resume next
	Dim WshShell
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.Run "Wpeutil Shutdown", 0, false
	Set WshShell = Nothing
	self.close	
End Sub

'Reboots WinPE enviroment
Sub Reboot()
	on error resume next
	Dim WshShell
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.Run "Wpeutil Reboot", 0, true
	Set WshShell = Nothing
	self.close
End Sub

'Executes a Command Shell to open
Sub CmdConsole()
	on error resume next
 	Dim WshShell
	Set WshShell = CreateObject("WScript.Shell")
 	WshShell.Run "cmd.exe", 1, true
 	Set WshShell = Nothing	
End Sub


'Gets a file name from a root path and a search pattern.comma seperated list
Function GetFileName(folderPath,strPattern)
		s = "FILE NOT FOUND:" & folderPath & " Pat:" & strPattern
		on error resume next
		Set RegEx = New RegExp
		RegEx.IgnoreCase = True
		RegEx.Global=True
		RegEx.Pattern=strPattern
    	Set fso = CreateObject("Scripting.FileSystemObject")
		Set fFolder = fso.GetFolder(folderPath)
		Set fFiles = fFolder.Files
		if fFiles.count > 1 then
			s=""
			'MsgBox "X" & fFiles.Item("name")
			For Each aFile in fFiles
				if RegEx.Test(aFile.name) Then
					s = s & aFile.name 
					s = s &  ","
				end if				
			Next
			s = Left(s,Len(s)-1)
		elseif fFile.count = 1 then
			s=""
				For Each aFile in fFiles
					if RegEx.Test(aFile.name) Then
						s = aFile.name 
					end if				
				Next				
		end if	
  	GetFileName = LTrim(RTrim(s))
End Function

'Returns free space in GB for specified drive
Function GetFreeSpace(strDrive)
	return = 0
	on error resume next
	Const HARD_DISK = 3
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DeviceID='" & strDrive & "'")
	For Each objDisk in colDisks
		return = Round(objDisk.FreeSpace / 1024 / 1024 /1024)
	Next
	GetFreeSpace = return
End Function

'Counts Files Within a Folder returns a number
Function GetFileCount(strFolderPath)
    counter = 0  
    on error resume next
    Set objFSO = CreateObject("Scripting.FileSystemObject")  
    Set objFolder = objFSO.GetFolder(strFolderPath) 
   ' Wscript.echo "PreCount:" & objFolder.Files.count
    For Each objFile In objFolder.Files 
        counter = counter + 1   
    Next
    GetFileCount=counter 
End Function

'Gets the date modified of a file
Function GetFileDate(strFile)
	return = Null
	On Error Resume Next
	SET objFSO = CreateObject("Scripting.FileSystemObject")
	SET objFile = objFSO.GetFile(strFile)
	return = CDate(objFile.DateLastModified)
	GetFileDate = return
End Function

'Compares a file to a  date in dd/mm/yyyy format. if file is greater than or equal is true. 
Function CompareFileDate(strFilePath,strDate)
	return = false
	On Error Resume Next
	SET objFSO = CreateObject("Scripting.FileSystemObject")
	SET objFile = objFSO.GetFile(strFilePath)
	fileDate = CDate(objFile.DateLastModified)
	compareDate = CDate(strDate)
	If fileDate >= compareDate then
		return = true
	else 
		return = false
	end if
	CompareFileDate = return
End Function

'Returns the file version of a file
function GetFileVersion(strFile)
	return = "NO_FILE_VERSION"
	On Error Resume Next
	Set objFSO=CreateObject("Scripting.FileSystemObject")
	return = objFSO.GetFileVersion(strFile)
    GetFileVersion = return
end function

'Moves a folder from one location to another on same partition
Function MoveFolder(strSource,strDest)
   return = "INIT"
   On Error Resume Next
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   fso.MoveFolder strSource,strDest
	if err.Number <> 0 then
		return = "MoveFolderFail:" & err.Number & " " & err.Description & " srcPath:" & strSource & " dstPath:" & strDest
	else
		return = "SUCCESS"
	end if
   MoveFolder=return
End function

'Get the current recovery drive partition drive letter
Function GetRecoveryDrive()
	SQL="SELECT DriveLetter FROM Win32_Volume Where Label = 'Recovery'"
	return = ""
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery(SQL)
	For Each objItem in colItems
		return = objItem.DriveLetter
	Next
	if return = "" then
		return = "NONE"
	end if
	GetRecoveryDrive = return
End Function
