'Production Setup Script for EFEC Lab and Media Kits
'========================== USER VARIABLES ===========================
'path for windows startup files ALL USERS
'winStartup = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"	

'Determines whether or not to use GUI
useGUI = true
oemInfoKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation"

'Current State Holder
msgFile = GetCurDir() & "NewGUIv2.0\msg.txt"

'DB tables
MasterCSV = GetCurDir() & "db\EFECmaster.csv"
LogFile = "c:\temp\Log.txt"

'Set's path from where to extract the BGInfo generated Desktop image
BMPPath = GetCurDir & "bin\bginfo\screen\BGInfo.bmp"
BgInfoEXE = GetCurDir() & "bin\Bginfo\BgInfo.exe"

'Sets the name of the data file to be used with script set'
BgInfMachineNameFile = GetCurDir() & "bin\bginfo\BGIdata\efec_systemname.txt"

'Product Code for Office required for activation
OffProdID = "NH3P2-CGFW2-Q7JXH-BGDYM-QJYG7" 

oemLogoLoc = GetCurDir() & "bin\images\oemlogo.bmp" 
oemLogoDest = "C:\Windows\system32\oemlogo.bmp"
strOSFDat = GetCurDir() & "bin\OSFDatNew\userinfov5.dat"
strOSFLoc = "C:\ProgramData\PassMark\OSForensics\"
strOSFOld = "C:\ProgramData\PassMark\OSForensics\userinfov5.dat"
strOSFNew = "C:\ProgramData\PassMark\OSForensics\olduserinfo.dat"

'  strAccountIcon = GetCurDir() & "bin\images\gear.png"
'  strAccountDestination = "C:\ProgramData\Microsoft\User Account Pictures"

'=========================ENVIROMENTAL VARIABLE ======================
'Creates a windows shell object for general use
Set WinShell = wscript.Createobject("wscript.Shell")

'vars to read help desk and build data
strBuild = ReadFile(GetCurDir() & "bin\bginfo\BGIdata\efec_BuildVersion.txt")
strIAVersion = ReadFile(GetCurDir() & "bin\bginfo\BGIdata\efec_IAVersion.txt")
helptxt = ReadFile(GetCurDir() & "bin\bginfo\BGIdata\efec_Contacts.txt")
BgiInfoProfile = " " & GetCurDir & "bin\Bginfo\efec.bgi /silent /timer=0"
if (useGUI) Then
	CreateAFile msgFile,"STARTING"
end if

'START HTA GUI	
If (useGUI) Then
	 	WinShell.Run "c:\windows\system32\mshta.exe D:\EFECECP007\dataset\setup\NewGUIv2.0\updateGUI.hta",1,False
		sendLog logFile,"HTA GUI has been launched. ","HTA"
		Wscript.Sleep 5000
		CreateAFile msgFile,"INIT"
		sendLog logfile, "Use GUI enabled, will proceed with GUI utilization", "GUI_Enabled"
	Else
		sendLog logfile, "Use GUI disabled, will proceed with running script.", "GUI_Disabled"
End If

'create temp folder for the log
Set fso = CreateObject("Scripting.FileSystemObject")
if not fso.FolderExists("C:\temp") Then
	fso.CreateFolder "C:\temp"
End If



'Start off the log file	
sendLog logFile,"start log Build: " & strBuild & vbCrlf & "------------------- SUPPORT CONTACTS----------------" & vbCrlf & helptxt & vbCrlf & vbCrlf,"STARTLOG"

'Define Error trap and global flags
errorFlag = False 'this sets the default Error flag for trapping all the steps. The sendLog function can set this.

'Get current MB Manufacturer
curMFG = MBMFG()
sendLog logFile,"MotherBoard MFG: [" & curMFG & "]","MFG"

'Get MotherBoard Serial Number
curSerialNumber = MBSerial() 
sendLog logFile, "Live System Serial Number: [" & curSerialNumber & "]", "MBSERIAL"

'Insures BGinfo loads on a new system; installs registry settings
winshell.RegWrite "HKEY_CURRENT_USER\Software\Sysinternals\BGInfo\EulaAccepted",1,"REG_DWORD"

'Ensure System is an EFEC system before executing script. Kill if not
if systemType = "NOT-EFEC" Then
	MsgBox "This is not an EFEC-MC System. Setup Canceled."
sendLog logFile,"This is not an EFEC-MC System. Setup Canceled : [" & systemType & " System]","SYSTEM_TYPE"
	Wscript.quit
End If	

' 'Create bin folder and copy contents from database
' if not fso.FolderExists("C:\dataset") Then
' 	fso.CreateFolder("C:\dataset")
' End If
' if not fso.FolderExists("C:\dataset\setup") Then
' 	fso.CreateFolder("C:\dataset\setup")
' End If
' if not fso.FolderExists("C:\dataset\setup\bin") Then
' 	fso.CreateFolder("C:\dataset\setup\bin")
' End If
' fso.CopyFolder GetCurDir() & "bin", "C:\dataset\setup\bin"
' fso.CopyFolder GetCurDir() & "NewGUI", "C:\dataset\setup"

' 'Change Account Icon
' iconDelete =  RenameFile("C:\ProgramData\Microsoft\User Account Pictures\user.png", "C:\ProgramData\Microsoft\User Account Pictures\olduser.png")
' if iconRename = "SUCCESS" then
'     sendLog logFile,"Account Icon Renamed","ICON_RENAME"
' Else
'     errorFlag = True
'     sendLog logFile,"FAILED to rename Account Icon.","FAIL_ICON_RENAME"
' End if

' iconCopy = MoveFile(strAccountIcon, strAccountDestination)
' if iconCopy = "SUCCESS" then
'     sendLog logFile,"Account Icon Copied","ICON_COPY"
' Else
'     errorFlag = True
'     sendLog logFile,"FAILED to copy Account Icon.","FAIL_ICON_COPY"
' End if

' iconRename = RenameFile("C:\ProgramData\Microsoft\User Account Pictures\gear.png", "C:\ProgramData\Microsoft\User Account Pictures\user.png")
' if iconRename = "SUCCESS" then
'     sendLog logFile,"Account Icon Changed","ICON_CHANGE"
' Else
'     errorFlag = True
'     sendLog logFile,"FAILED to change Account Icon.","FAIL_ICON_CHANGE"
' End if

'======================= START SCRIPT EXECUTION ======================
'Set state to INIT
'if (useGUI) then

'TRAP THE LAUNCHING OF THE HTA GUI UNTIL SCREEN RESOLUTION IS READY. THIS ENSURES SYSTEM IS READ BEFORE GUI LAUNCHES
	' check current resolution of system
' 	curScreenRes = ScreenWidth()
' 	startTimer = Timer()
' 	timeResult = 0
' 		sendLog logFile,"Starting Resolution Check for Launching HTA GUI: CurRes: [" & curScreenRes & "]","ScreenCheck"
		
' 'ENSURES SCREEN RESOLTUION MEETS MIN FOR EACH DEVICE BEFORE LAUNCHING GUI TIMES OUT AFTER TIMER LIMIT
'   Do Until (curScreenRes >= minScreenRes) OR (timeResult > 180)
' 		timeResult = FormatNumber(Timer() - StartTime, 2)
' 		Wscript.Sleep 5000
' 		curScreenRes = ScreenWidth()
' 	Loop
		
' 	'revalidate before launching HTA
' 	If  (curScreenRes >= minScreenRes) Then 
' 		sendLog logFile,"Passed Screen Res Validation. Starting HTA GUI. CurRes: [" & curScreenRes & "]","ScreenCheck"
' 	Else
' 			sendLog logFile,"Screen Resolution Failed but continuing setup.CurRes: [" & curScreenRes & "]","ScreenCheck"
' 	End If
'End If

If (useGUI) Then
	'Wait until GUI HTA releases loop.
	Do Until (readMsg ="INIT") OR (readMsg ="IN-PROGRESS")
		readMsg = ReadFile(msgFile)
		Wscript.Sleep(2000)
	'Break loop if HTA is no longer there
		if IsCMDRunning("updateGUI.hta") = False Then
		sendLog logFile,"GUI is no longer running. closed script","GUI_CLOSED"
		wscript.quit
		End if
	Loop
End If

'NOT CURRENTLY IMPLEMENTED, PLACEHOLDER FOR FUTURE DEVELOPMENT
'SET NEW IP ADDRESS
' If curMAC <> "" Then
'     'extract decimals octets from MAC address for creating IP address
'     newIP = "192.168." & CLng("&h" & Mid(curMAC,13,2)) & "." & CLng("&h" & Right(curMAC,2))
'     SetIp(newIP)
'     sendLog logFile, "New Network IP Address Is Set: [" & newIP & "]", "IP" 
' End If

'SETS SYSTEM TIME ZONE TO UTC
If (useGUI) Then
	Call GUICheck()
	'Set State UTC
	CreateAFile msgFile,"START_UTC_TIMEZONE"
	Wscript.Sleep(1000)
End If

RunApp	"tzutil /s UTC",True
	sendLog LogFile,"Time Zone set to UTC (GMT)","TIMEZONE"
	If (useGUI) Then
		CreateAFile msgFile, "PASS_UTC_TIMEZONE"
		Wscript.Sleep(1000)
	End If

'NOT CURRENTLY IMPELEMENTED
' 'attempt to sync time with a server	
' pingResult = ReturnPing(serverIP)
' if  pingResult = True Then
'     sendLog LogFile,"Attempting to Set Time.","TIMESYNC"
'     valSetTime = SetTimeDate(serverIP)
'         if valSetTime = 0 Then
'             sendLog LogFile,"Time Synchronized with a Server:[" & serverIP & "].","TIMESYNC"
'         else
'             sendLog LogFile,"Time Sync failed with a Server:[" & serverIP & "] with error: " & valSetTime,"FAIL_TIMESYNC"
'         end if
'     'validate the time has been set
'     valTimeDate = GetTimeMatch(serverIP)
'     if valTimeDate = "PASS" Then
'         sendLog LogFile,"Time Validation has passed the test" ,"TIME_CHECK"
'     Else
'         sendLog LogFile,"FAILED Time Validation. Set to UTC Manually.","TIME_SET"
'         ReplaceFile "msg.txt","IN-PROGRESS-UTC"
'         wscript.sleep 15000
'         ReplaceFile "msg.txt","IN-PROGRESS"
'     End If
' Else
'     sendLog LogFile,"Manually Validate and Set UTC Time on System.","TIME_SET"
'     ReplaceFile "msg.txt","IN-PROGRESS-UTC"
'     wscript.sleep 15000
'     ReplaceFile "msg.txt","IN-PROGRESS"
' End If

If (useGUI) Then
	'Run OSForensics AutoIt Script
	Call GUICheck()
	CreateAFile msgFile, "START_OSF_REG"
End If

'RunApp GetCurDir() & "OSForensicsScript.exe", True
Call RenameFile(strOSFOld, strOSFNew)
Call MoveFile(strOSFDat, strOSFLoc)
If (FileExists(strOSFDat)) Then
	If (useGUI) Then
		CreateAFile msgFile, "PASS_OSF_REG"
	End If
	sendLog LogFile, "OSForensics Registered", "OSForensics"
else 
	If (useGUI) Then
		CreateAFile msgFile, "FAIL_OSF_REG"
	End If
	sendLog LogFile, "OSForensics Failed to Register", "OSFFail"
End If

If (useGUI) Then
	'SYSTEM NAME CHANGE 
	Call GUICheck()
	CreateAFile msgFile, "START_SET_COMP_NAME"
End If

sysModel = GetCPUName()
    sendLog LogFile, "Ck SysName " & sysModel, "GetModel"

	strModel = sysModel

	'System Name Prefixes
		''pfxSEEK = "" 'usually no prefix
			''pfxDELL = "DELL-" 
			pfx = ""

Select Case sysModel

case "PRECISION7520"
   systemType = "LAB"
   'pfx = "LAB-"
   minScreenRes = 1366
   ChangeComputerName("LAB-MASTER")

case "LAB-MASTER"
	systemType = "LAB"
	'pfx = "LAB-"
	minScreenRes = 1366

case "PRECISIONM6800"
   systemType = "NSFTC"
   'pfx = "NSFTC-"
   minScreenRes = 1366
   ChangeComputerName("NSFTC-MASTER")

case "NSFTC-MASTER"
	systemType = "NSFTC"
	'pfx = "NSFTC-"
	minScreenRes = 1366

case "FZM1"
   systemType = "MEDIA"
   'pfx = "MEDIA-"
	minScreenRes = 1024
	ChangeComputerName("MEDIA-MASTER")

case "MEDIA-MASTER"
	systemType = "MEDIA"
	'pfx = "MEDIA-"
	minScreenRes = 1024

case else systemType = "NOT-EFEC"
	If (useGUI) Then
		CreateAFile msgFile, "FAIL_SET_COMP_NAME"
	End If
	sendLog logfile, "Failed to set computer name."
end select 

CPUName = systemType & "-" & curSerialNumber
'wscript.echo "CPUName: " & CPUName

  If (useGUI) Then
	CreateAFile msgFile, "PASS_SET_COMP_NAME"
  End If
  sendLog logFile,"System Type Discovered to be : [" & systemType & " System]","SYSTEM_TYPE"
	sendLog logFile, "Live System Serial Number: [" & curSerialNumber & "]", "MBSERIAL"
	sendLog logFile, "CPU Name: [" & CPUName & "]", "MBSERIAL"

IF (useGUI) Then
	'SET  BACKGROUND DESKTOP IMAGE
	Call GUICheck()
	CreateAFile msgFile, "START_SET_BACKGROUND"
End If

'write the SystemName text file to the BGInfo data sources	
Call DeleteFile(BgInfMachineNameFile)	
Call CreateAFile(BgInfMachineNameFile,newMachineName)

'Run Binfo with Profile to generate new screen image
WinShell.run BgInfoEXE & BgiInfoProfile,0,true
    sendLog logFile,"New Desktop Background has been generated.","BGINFO_RUN"
        
'move BGinfo bmp to Repository and move OEM Logo
BMPDest = "c:\windows\BGInfo.bmp"
bmpCopy = CopyFile(BMPPath,BMPDest)
oemCopy = CopyFile(oemlogoloc, oemLogoDest)
if bmpCopy = "SUCCESS" then
	sendLog logFile,"Desktop Background has been activated on the system.","DESKTOP_COPY"
	If (useGUI) Then
		Wscript.Sleep(1000)
		CreateAFile msgFile, "PASS_SET_BACKGROUND"
		Wscript.Sleep(1000)
	End If
Else
    errorFlag = True
	sendLog logFile,"FAILED to copy Desktop background into Windows default directory.","FAIL_DESKTOP_COPY"
	If (useGUI) Then
		Wscript.Sleep(1000)
		CreateAFile msgFile, "FAIL_SET_BACKGROUND"
		Wscript.Sleep(1000)
	End If
End if

if oemCopy = "SUCCESS" then
    sendLog logFile,"Oem Logo has been relocated to system folder.","OEM_COPY"
Else
    errorFlag = True
    sendLog logFile,"FAILED to copy OEM Logo into Windows default directory.","FAIL_OEM_COPY"
End if

If (useGUI) Then
	'WINDOWS MAK ACTIVATION	
	Call GUICheck()
	CreateAFile msgFile, "START_ACTIVATE_WINDOWS"
End If

'Get Current Actual InstallID from live system.
curMAK_Install = GetWinID()
    sendLog logFile,"Current Win InstallID: [" & curMAK_Install & "]", "curMAK_Install"
'Find Current InstallID from DB and trim out special chars
newMAK_Install = Replace(GetFieldFromProduct(MasterCSV,curSerialNumber, "Win10E", 2),"_","")
	sendLog logFile,"Lookup Win InstallID: [" & newMAK_Install  & "]", "newMAK_Install"	
'Get Current Windows ACTIVATION status
    curActivation = isWinAct()
        sendLog logFile,"Current Windows Activation Status: [" & curActivation & "]","WIN_STATUS"
    'Compare and live vs lookup InstallID pass match from existing vs lookup InstallID
        If newMAK_Install = curMAK_Install Then
        'lookup confID to activate windows with
            confID = Replace(GetFieldFromProduct(MasterCSV,curSerialNumber, "Win10E",3),"_","")
                if confID <> "" Then
                    sendLog logFile,"Matching Windows ConfID: [" & confID & "]", "confID"
                Else
                    errorFlag = True
                    sendLog logFile,"FAILED Lookup of Confirmation ID: Contact Support Center.","FAIL_CONFID" 
                End If
                
            'Make sure windows isn't already activated then activate windows		
            if curActivation = False Then
                
                activateWin = ActivateWindows(confID)
                    
                    'validate windows activation after attempted activation
                    If activateWin = "SUCCESS" Then
						sendLog logFile, "Successfully activated Windows", "Activation"
						If (useGUI) Then
							Wscript.Sleep(1000)
							CreateAFile msgFile, "PASS_ACTIVATE_WINDOWS"
							Wscript.Sleep(1000)
						End If
                    Else
                        errorFlag = True
						sendLog logFile, "FAILED Win Activation with stored ConfID. Contact Support Center: " & activateWin,"FAIL_ACTIVATION"
						If (useGUI) Then
							Wscript.Sleep(1000)
							CreateAFile msgFile, "FAIL_ACTIVATE_WINDOWS"
							Wscript.sleep(1000)
						End If
                    End If
            Else
                'Windows already Activated
				sendLog logFile, "Windows was previously already Activated.","Activation"
				If (useGUI) Then
					Wscript.Sleep(1000)
					CreateAFile msgFile, "FAIL_ACTIVATE_WINDOWS"
					Wscript.sleep(1000)
				End If
            End If
                
        Else
            'failed look up last chance to see if windows is already activate
            
            if curActivation= False Then
                errorFlag = True
				sendLog logFile,"FAILED WinActivation. Mismatched InstallID. Contact the Support Center", "FAIL_MAK"
				If (useGUI) Then
					Wscript.sleep(1000)
					CreateAFile msgFile, "FAIL_ACTIVATE_WINDOWS"
					Wscript.sleep(1000)
				End If
            Else
                errorFlag = True
				sendLog logFile,"Mismatched MAK but Windows Already Activated. Contact the Support Center.", "FAIL_ACTIVATION"
				If (useGUI) Then
					Wscript.sleep(1000)
					CreateAFile msgFile, "FAIL_ACTIVATE_WINDOWS"
					Wscript.sleep(1000)
				End If
            End If
        End If	

If (useGUI) Then
	'Office Activation
	Call GUICheck()
	CreateAFile msgFile, "START_ACTIVATE_OFFICE"
End If
'Activate Office License if neccessary
isOfficeLicensed()
'Get Current Office ACTIVATION status
curActivation = isOfficeAct()
'Make sure office isn't already activated then activate office		
if curActivation = False Then
	OfficeConfID = Replace(GetFieldFromProduct(MasterCSV,curSerialNumber, "Office2013",3),"_","")
                if confID <> "" Then
                    sendLog logFile,"Matching Office ConfID: [" & OfficeConfID & "]", "OfficeConfID"
                Else
                    errorFlag = True
					sendLog logFile,"FAILED Lookup of Confirmation ID: Contact Support Center.","FAIL_OfficeCONFID"
					IF (useGUI) Then
						Wscript.sleep(1000)
						CreateAFile msgFile, "FAIL_ACTIVATE_OFFICE" 
						Wscript.sleep(1000)
					End If
                End If
	activateOff = ActivateOffice(OfficeConfID)	
    'validate office activation after attempted activation
    If activateOff = "SUCCESS" Then
		sendLog LogFile, "Successfully activated Office", "SUCCCESS_ACTIVATION_OFFICE"
		If (useGUI) Then
			Wscript.Sleep(1000)
			CreateAFile msgFile, "PASS_ACTIVATE_OFFICE"
			Wscript.Sleep(1000)
		End If
        'objStdOut.Write "Successfully activated Office" & vbNewLine
    Else
        errorFlag = True
		sendLog LogFile, "FAILED Office Activation with stored ConfID:" & activateOff, "FAIL_ACTIVATION_OFFICE"
		If (useGUI) Then
			Wscript.sleep(1000)
			CreateAFile msgFile, "FAIL_ACTIVATE_OFFICE"
			Wscript.sleep(1000)
		End If
        'objStdOut.Write "FAILED Office Activation with stored ConfID:" & activateOff & vbNewLine
    End If
Else
	'Office already Activated
	errorFlag = True
	sendLog LogFile, "Office has already been Activated.", "ACTIVATION_OFFICE"
	If (useGUI) Then
		Wscript.sleep(1000)
		CreateAFile msgFile, "FAIL_ACTIVATE_OFFICE"
		Wscript.sleep(1000)
	End If
	'objStdOut.Write "Office has already been Activated." & vbNewLine
End If

'NOT SURE IF WE NEED THIS OR WILL IN THE FUTURE
' 'SETUP Technical Manual ICON LINK
' workdir = "c:\dataset\docs"
' targetShorcut = "c:\users\public\Desktop\"
' targetName = "User Manual"
'     Call CreateShortcut(techManPath,workdir,targetShorcut,targetName)	

'NEED TO REVIEW THIS
'IMPORT REG KEYS FOR Product Registry settings for Windows 10
If (useGUI) Then
	'Office Activation
	Call GUICheck()
	CreateAFile msgFile, "START_IMPORT_REG_KEYS"
End If

strBuild = Clean(strBuild)
mtchIAVersion = Clean(mtchIAVersion)
helpSet = Split(helptxt,vbLf)

SupportHours = "24/7 hrs. " & Clean(helpSet(1)) 
SupportPhone = Clean(HelpSet(0))
Model = systemType & " Build: " & strBuild 

	rt1 = setReg(oemInfoKey,"Model",Model,False)
	rt2 = setReg(oemInfoKey,"SupportHours",SupportHours,False)
	rt3 = setReg(oemInfoKey,"SupportPhone",SupportPhone,False)
	rt4 = setReg(oemInfoKey,"Logo","c:\\windows\\system32\\oemlogo.bmp",False)
	rt5 = setReg(oemInfoKey,"Manufacturer","EFEC - Marine Corps",False)
	rt6 = setReg(oemInfoKey,"SupportURL","",False)

	if rt1+rt2+rt3 = 0 Then
		sendLog LogFile,"OEMInfo Registry updated successfully","PRODUCT"	
		if (useGUI) Then
			CreateAFile msgFile, "PASS_IMPORT_REG_KEYS"
		End if
	else
		sendLog LogFile,"FAIL OEMInfo update: RT1:" & rt1 & " RT2:" & rt2  & " RT3:" & rt3,"FAIL_PRODUCT"
		if (useGUI) Then
			CreateAFile msgFile, "FAIL_IMPORT_REG_KEYS"
		end if
	end if		

if errorFlag Then
	sendLog logFile,"----------------ENDED WITH FAIL ERRORS---------------------", "ERRORFLAG"
	'objStdOut.Write "----------------ENDED WITH FAIL ERRORS---------------------"
else
	sendLog logFile,"----------------------NO ERRORS--------------------------", "NO_ERRORFLAG"
	'objStdOut.Write "----------------------NO ERRORS--------------------------"
End If

Wscript.Echo "Script has completed."
' =========================FUNCTIONS=====================================
'Deletes a file or files
Function DeleteFile(strFile)
	On Error Resume Next
	dim objFSO
	set objFSO=CreateObject("Scripting.FileSystemObject")
	'to force deletion, set to TRUE otherwise set to FALSE
	objFSO.DeleteFile strFile,TRUE
End Function

' 'read SerialNumber from file matching MACAddress assumes MAC is first field
' Function GetFieldFromSerial(strFile,strSerialNumber,FieldPosition)
' 	return = "" 'Returns NONE if match not found
' 	On Error Resume Next
' 	Const OpenAsDefault = -2
' 	Const FailIfNotExist = 0
' 	Const ForReading = 1
' 	Const ForAppending = 8
' 	FilePath = strFile
' 	Set oFSO = CreateObject("Scripting.FileSystemObject")
' 	If oFSO.FileExists(FilePath) Then
' 		Set fFile = oFSO.OpenTextFile(FilePath,ForReading,FailIfNotExist, OpenAsDefault)
' 	Else
' 		MSGBOX "Unable to find File to Read. Check the path" & vblf & FilePath	
' 		Set oFSO = Nothing
' 		Exit Function
' 	End If
' 	sResults = fFile.ReadAll
' 	fFile.Close
' 	 'Wscript.Echo "file:" & sResults
' 	SplitLine = Split(sResults, vbLF)
' 	RowLen = Ubound(SplitLine)
' 		'Wscript.Echo RowLen
' 		For ii = 0 to RowLen
' 			SplitData = Split(SplitLine(ii),",")
' 				'Wscript.Echo SplitData(0)
' 				If SplitData(0) = strSerialNumber Then
' 					'Wscript.Echo "Found it. ParentIP is: " & SplitData(2)
' 					return = SplitData(FieldPosition)
' 				End If
' 		Next
' 		return = Trim(Replace(Replace(return,vbCrLf,""),vbTab,""))
' 		GetFieldFromSerial = return
' End Function

Function GetFieldFromProduct(strFile, strSerialNumber, strProduct, FieldPosition)
	return = ""
	On Error Resume Next
	Const OpenAsDefault = -2
	Const FailIfNotExist = 0
	Const ForReading = 1
	Const ForAppending = 8
	FilePath = strFile
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(FilePath) Then
		Set fFile = oFSO.OpenTextFile(FilePath,ForReading,FailIfNotExist, OpenAsDefault)
	Else
		MSGBOX "Unable to find File to Read. Check the path" & vblf & FilePath	
		Set oFSO = Nothing
		Exit Function
	End If
	sResults = fFile.ReadAll
	fFile.Close
	SplitLine = Split(sResults, vbLF)
	RowLen = Ubound(SplitLine)
	For ii = 0 to RowLen
		SplitData = Split(SplitLine(ii),",")
		if (SplitData(0) = strSerialNumber) and (SplitData(1) = strProduct) Then
			return = SplitData(FieldPosition)
		End If
	Next
	return = Trim(Replace(Replace(return,vbCrLf,""),vbTab,""))
		GetFieldFromProduct = return
End Function

'NOT CURRENTLY USED
'Gets current model name of system
' Function GetModel()
' 	return= ""
' 	On Error Resume Next
' 	Const wbemFlagReturnImmediately = &h10
' 	Const wbemFlagForwardOnly = &h20
' 	Set objWMIService = GetObject("winmgmts:\\.\ROOT\CIMV2")
' 	strQuery = "SELECT * FROM Win32_ComputerSystem"
' 	Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
' 		For Each objItem in colItems
' 		   return = objItem.Model
' 		Next
' 	GetModel = Replace(return," ","")
' End Function

'Get Computer Name from reg
Function GetCPUName()
regComputerName = "HKLM\SYSTEM\CurrentControlSet\Control\" & "ComputerName\ComputerName\ComputerName"
Set objShell = CreateObject("WScript.Shell")
GetCPUName = objShell.RegRead(regComputerName)
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Changes HostName on a System. returns FAIL, SUCCESS, DENIED_ACCESS, ERROR_CODE
Function ChangeComputerName(strNewName)
	ChangeComputerName = "FAIL"
	Set NetObj = wscript.Createobject("wscript.network")
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	set objInst = objWMI.InstancesOf("Win32_ComputerSystem")
	For Each obj in objInst
		err = obj.Rename(strNewName)
		If err = 0 then
			ChangeComputerName = "SUCCESS"
			
		ElseIf err = 5 Then
			ChangeComputerName = "DENIED_ACCESS" 'Access. Please check whether UAC is activated or you have admin priveleges."
		Else
			ChangeComputerName = "ERROR_CODE: " & err
		End if
	Next
End Function

'Activate Windows with a ConfirmationID
Function ActivateWindows(strConfID)
	ActivateWindows = "INIT"
	'Product Code for Windows10E required for activation
	winActID = "2ffd8952-423e-4903-b993-72a1aa44cf82"
	'Send output to logfile
	'output error and detect success with RegEx
 	binFile = "cscript c:\windows\system32\slmgr.vbs /atp "
	Set oShell = CreateObject("Wscript.Shell")
	tmpResult = GetBINResult(binFile & strConfID & " " & winActID)
	if REGEX_GetMatch(tmpResult,"deposited") = "deposited" then
		ActivateWindows = "SUCCESS"
	else
		ActivateWindows = tmpResult
	end if
End Function
	
'Get the Windows Installation ID for offline Windows Activation
Function GetWinID()
	binFile = "cscript c:\windows\system32\slmgr.vbs /dti"
	strPattern = "(\d{63})"
	strResult = GetBINResult(binFile)
	GetWinID = REGEX_GetMatch(strResult,strPattern)
End Function

'check to see if windows is currently activated
Function isWinAct()
	return = False
	On Error Resume Next
	binFile = "cscript c:\windows\system32\slmgr.vbs /dlv"
	strResult = GetBINResult(binFile)
		newResult = REGEX_GetMatch(strResult,"(Licensed)")
		If newResult = "" Then
			return = False
		Else
			return = True
		End IF
	isWinAct = return
End Function

'Gets current operating directory
Function GetCurDir()
On Error Resume Next
	GetCurDir=Left(WScript.ScriptFullName,Len(WScript.ScriptFullName)_
	-Len(WScript.ScriptName))
End Function

''''''''''''''''''''''''''''''''''''''''
'reads a file's contents into memory
Function ReadFile(filePath)
	Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
	ReadFile = oFileSystemObject.OpenTextFile(filePath,1).ReadAll
	
End Function

''''''''''''''''''''''''''''''''''''''''
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
		CopyFile = "FAIL:Err:" & err.number
	end If
End Function

'Create a file with specic content
Function CreateAFile(strFileName,strContent)
	Const AppendIfExists = False
	Const ForReading = 1
	Const ForAppending = 8
	Const ForWriting = 2
	on error resume next
	Set FS=CreateObject("Scripting.FileSystemObject")
	Set fFile = FS.OpenTextFile(strFileName, ForWriting, AppendIfExists)
	fFile.Write(strContent)
	fFile.Close
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

Sub RunApp(appName,waitState)
    Set objShell = CreateObject("Wscript.Shell")
	objShell.Run appName,2,waitState
End Sub

'Activate office with a ConfirmationID
Function ActivateOffice(strOfficeConfID)
	ActivateOffice = "INIT"
	'Send output to logfile
	'output error and detect success with RegEx
	Set fso = CreateObject("Scripting.FileSystemObject")
	newfolder = GetFolderName("c:\Program Files\Microsoft Office\", "(Office)(\d{2})")
	file32 = "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs"
	file64 = "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs"
	if (FileExists(file32)) then
		binFile = "cscript " & chr(34) & "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /actcid:"
	End If
	if (FileExists(file64)) then
		binFile = "cscript " & chr(34) & "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /actcid:"
	End If
	Set oShell = CreateObject("Wscript.Shell")
	tmpResult = GetBINResult(binFile & strOfficeConfID)
	if REGEX_GetMatch(tmpResult,"(successful)") = "successful" then
		ActivateOffice = "SUCCESS"
	else
		ActivateOffice = tmpResult
	end if
End Function
	
'Get the Office Installation ID for offline Office Activation
Function GetOfficeID()
	Set fso = CreateObject("Scripting.FileSystemObject")
	newfolder = GetFolderName("c:\Program Files\Microsoft Office\", "(Office)(\d{2})")
	file32 = "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs"
	file64 = "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs"
	if (FileExists(file32)) then
		binFile = "cscript " & chr(34) & "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /dinstid"
	End If
	if (FileExists(file64)) then
		binFile = "cscript " & chr(34) & "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /dinstid"
	End If
    strPattern = "(\d{63})"
	strResult = GetBINResult(binFile)
	GetOfficeID = REGEX_GetMatch(strResult,strPattern)
End Function

'I BELIEVE THIS IS REDUNTANT
' 'Retrieve Office Confirmation ID
' Function GetConfID()
' 	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(OfficeTXT,1)
' 	strResult = objFileToRead.ReadAll()
' 	strPattern = "(\d{48})"
' 	GetConfID = REGEX_GetMatch(strResult, strPattern)
' 	objFileToRead.Close
' 	Set objFileToRead = Nothing
' End Function

'Attempts to pull Serial number
Function MBSerial() 
	Set WMI = GetObject("WinMgmts:")
	Set objs = WMI.InstancesOf("Win32_Bios")
	For Each obj In objs
  		sAns = sAns & obj.SerialNumber
 	If sAns < objs.Count Then sAns = sAns & ","
	Next
	MBSerial = sAns
End Function

'Get results from a command shell
Function GetBINResult(binFile)
	Set return = Wscript.StdIn
	on error resume next
	Dim oShellOutputFileToRead, iErr
	Set oShellObject = CreateObject("Wscript.Shell")
	Set oFileSystemObject = CreateObject("Scripting.FileSystemObject")
	Set fso = CreateObject("Scripting.FileSystemObject")
	' path="C:\temp\"   
	' exists = fso.FolderExists(path)
	' if NOT (exists) then 
	' 	fso.CreateFolder "C:\temp"
	' End If
	sShellRndTmpFile = "c:\temp\cmd.temp"
	oShellObject.Run "%comspec% /c " & binFile & " > " & sShellRndTmpFile, 0, True
	iErr = Err.Number
   	On Error GoTo 0
	If iErr <> 0 Then 
		return = Err.Description
		Exit Function 
	End If 
	Wscript.Sleep 2500
	return = oFileSystemObject.OpenTextFile(sShellRndTmpFile,1).ReadAll
    oFileSystemObject.DeleteFile sShellRndTmpFile, True 
    GetBINResult = return
End Function

'Uses Regular Expressions to grab a match via a pattern
Function REGEX_GetMatch(strString,strPattern)
	'On Error resume next
	Dim RegEx
	Set RegEx = New RegExp
	RegEx.IgnoreCase = True
	RegEx.Global=False
	RegEx.Pattern=strPattern
	Set strFound = RegEx.Execute(strString)
	For Each key in strFound
		REGEX_GetMatch = key.value
	Next
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

Function GetFolderName(folderPath,strPattern)
	Dim RegEx
		Set RegEx = New RegExp
		RegEx.IgnoreCase = True
		RegEx.Global=True
		RegEx.Pattern=strPattern
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderPath)
    Set fc = f.SubFolders
    For Each f1 in fc
        s = s & f1.name 
        s = s &  vbCrLf
    Next
      sResults = s
    NewResults = Split(sResults,vbCrLF)
	LSTLEN = Ubound(NewResults)
	For ii = 0 to LSTLEN
		strSearchString = NewResults(ii)
		'Wscript.Echo "Line="& strSearchString
		Set colMatches = RegEx.Execute(strSearchString)
			'Wscript.Echo "Matches " & colMatches.Count
		If colMatches.Count > 0 Then
			For Each strMatch in colMatches
				si = NewResults(ii)
			Next
		End If
	Next
	GetFolderName = si
End Function

'check to see if Office is currently activated
Function isOfficeAct()
	return = False
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	newfolder = GetFolderName("c:\Program Files\Microsoft Office\", "(Office)(\d{2})")
	file32 = "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs"
	file64 = "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs"
	if (FileExists(file32)) then
		binFile = "cscript " & chr(34) & "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /dstatus"
	End If
	if (FileExists(file64)) then
		binFile = "cscript " & chr(34) & "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /dstatus"
	End If
	strResult = GetBINResult(binFile)
		newResult = REGEX_GetMatch(strResult,"(LICENSED)")
		If newResult = "" Then
			return = False
		Else
			return = True
		End IF
	isOfficeAct = return
End Function

Function isOfficeLicensed()
	return = False
	On Error Resume Next
	Set fso = CreateObject("Scripting.FileSystemObject")
	newfolder = GetFolderName("c:\Program Files\Microsoft Office\", "(Office)(\d{2})")
	file32 = "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs"
	file64 = "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs"
	if (FileExists(file32)) then
		binFile = "cscript " & chr(34) & "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /dstatus"
	End If
	if (FileExists(file64)) then
		binFile = "cscript " & chr(34) & "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /dstatus"
	End If
	strResult = GetBINResult(binFile)
		myResult = REGEX_GetMatch(strResult, "(No installed)")
		If myResult = "No installed" Then
			activateOfficeID()
		End If
	isOfficeAct = return
End Function

Function activateOfficeID()
	Set objShell = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")
	newfolder = GetFolderName("c:\Program Files\Microsoft Office\", "(Office)(\d{2})")
	file32 = "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs"
	file64 = "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs"
	if (FileExists(file32)) then
		objShell.run "cscript " & chr(34) & "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /inpkey:" & OffProdID, 0, True
		sendLog LogFile, "Sent Input Product Key for Office Activation"
	End If
	if (FileExists(file64)) then
		objShell.run "cscript " & chr(34) & "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /inpkey:" & OffProdID, 0, True
		sendLog LogFile, "Sent Input Product Key for Office Activation"
	End If
End Function

'Send a logFile message
Function sendLog(strLogfile,strMessage,strEvent)
	Const AppendIfExists = True
	Const ForReading = 1
	Const ForAppending = 8
	on error resume next
	Set FS=CreateObject("Scripting.FileSystemObject")
		Set fFile = FS.OpenTextFile(strLogfile,ForAppending,AppendIfExists)
	strContent = Date() & "." & Time() & "," & strMessage &	"," & strEvent & "$" 
		fFile.Write(strContent & vbCrlf)
	fFile.Close
	sendLog = strContent
	Wscript.Sleep 555
End Function

'NEED TO TEST THIS
'Sub that turns off UAC and Clears Autologin and restores security
Function EnableUAC()
	'removes startup icon from both systems
	DeleteFile(EnvPath("%USERPROFILE%") & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\setup.lnk")

	rt1 = DelReg("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultUserName")
	rt2 = DelReg("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","AutoAdminLogon")
	rt3 = DelReg("SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon","DefaultPassword")
	rt4 = setReg("SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","EnableLUA",1,True)
	rt5 = setReg("SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","ConsentPromptBehaviorAdmin",2,True)
	t5 =  setReg("SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","ConsentPromptBehaviorUser",1,True)
	rt6 = setReg("SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System","PromptOnSecureDesktop",1,True)
		if rt1+rt2+rt3+r+4 <> 0 Then
			return = "R1:" & rt1 & " R2:" & rt2 & " R3:" & rt3 & " R4:" & rt4
		else
			return = "GOOD"
		end if
	SetServiceState "EMET_Service","Automatic"
End Function


'Sets a service into a startup state. Disabled, Manual, or Automatic by name.
Function SetServiceState(strService,strState)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		Set colServiceList = objWMIService.ExecQuery _
	    ("SELECT * FROM Win32_Service WHERE Name like '%" & strService & "%'")
	For Each objService in colServiceList
	    errReturnCode = objService.Change( , , , , strState)   
	Next
End Function
'expands enviromental paths for windows
Function EnvPath(strDosVariable)
	return = ""
	Set wshShell = CreateObject( "WScript.Shell" )
	return = wshShell.ExpandEnvironmentStrings(strDosVariable)
	EnvPath = return
End Function
' Delete a Registry Key
Function DelReg(strPath,strValue)
	On Error Resume Next
	Const HKEY_LOCAL_MACHINE = &H80000002
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	 'strKeyPath = "SOFTWARE\System Admin Scripting Guide"
	 result = oReg.DeleteValue(HKEY_LOCAL_MACHINE,strPath,strValue)
	 DelReg = result
End Function
'Set change a registry key
Function setReg(strKeyPath,strKeyName,strKeyValue,boolDWORD)
	Const HKEY_LOCAL_MACHINE = &H80000002
	On Error Resume NExt
	Set oReg=GetObject("winmgmts://./root/default:StdRegProv")
	 'strKeyPath = "SOFTWARE\System Admin Scripting Guide"
	 'strKeyName = "String Value Name"
	 'strValue = "string value"
	if boolDWORD = True Then
		result = oReg.SetDWORDValue(HKEY_LOCAL_MACHINE,strKeyPath,strKeyName,strKeyValue)
	Else
		result = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strKeyName,strKeyValue)
	End If
		setReg = result	
End Function

'Check for GUI and kills current script if not found
Function GUICheck()
	if IsCMDRunning("updateGUI.hta") = False Then
		sendLog logFile,"GUI is no longer running. closed script","GUI_CLOSED"
		Wscript.Quit
	End if
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

'Clean out line feeds and Carriage returns
Function Clean(newString)
	return = Replace(newString,vbLf,"")
	clean = Replace(return,vbCr,"")
End Function

'MISSING SOMETHING FROM GETHOSTFROMMAC

'JAVA PERMISSIONS?

'USED FOR TM ICON CREATION
' 'Creates a short to a program or file. 
' Function CreateShortcut(strProgram, strWorkingDir,strShortcutPath,strShortcutName)
' 	Set objshell=CreateObject("WScript.Shell")
' 		set objShortcut = objshell.CreateShortcut(strShortCutPath & "\" & strShortCutName & ".lnk")
' 	objshortcut.TargetPath = strProgram
' 	objshortcut.WindowStyle = 1
' 		Description = InStrRev(strProgram,".")
' 		Description = Mid(strProgram,1,Description)
' 	objshortcut.Description = Description
' 	objshortcut.WorkingDirectory = strWorkingDir
' 	objshortcut.IconLocation= strProgram
' 	objshortcut.Hotkey = ""
' 	objshortcut.Save
' End Function

' 'Clean out line feeds and Carriage returns
' Function Clean(newString)
' 	return = Replace(newString,vbLf,"")
' 	clean = Replace(return,vbCr,"")
' End Function

'NEEDED FOR IP SYNC
' 'sets date time from a server
' Function SetTimeDate(strIPAddress)
' 	On Error Resume Next
' 	Const wbemFlagReturnImmediately = &h10
' 	Const wbemFlagForwardOnly = &h20
' 	Set wshNetwork = WScript.CreateObject("WScript.Network")
' 	strQuery = "SELECT * FROM Win32_OperatingSystem"
' 		Set objWMIRemote = GetObject("winmgmts:\\" & strIPAddress & "\ROOT\CIMV2")
' 		Set objWMILocal = GetObject("winmgmts:\\.\ROOT\CIMV2")
' 	'Exec Remote Query
' 	Set colItems = objWMIRemote.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
' 	For Each objItem in colItems
' 		return1 = objItem.LocalDateTime
' 	Next
' 		'Exec local query
' 	Set colItems = objWMILocal.ExecQuery(strQuery)
' 	For Each objItem in colItems
' 		return2 = objItem.SetDateTime(return1)
' 	Next
' 	SetTimeDate = return2
' End Function
' Function GetTimeMatch(strIPAddress)
' 	return = "FAIL:OTHER"
' 	On Error Resume Next
' 	Const wbemFlagReturnImmediately = &h10
' 	Const wbemFlagForwardOnly = &h20
' 	Set wshNetwork = WScript.CreateObject("WScript.Network")
' 	strComputer = strIPAddress
' 	strQuery = "SELECT * FROM Win32_UTCTime"
' 		Set objWMIRemote = GetObject("winmgmts:\\" & strComputer & "\ROOT\CIMV2")
' 		Set	objWMILocal = GetObject("winmgmts:\\.\ROOT\CIMV2")
' 		'Exec Remote Query
' 	Set colItems = objWMIRemote.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
' 	For Each objItem in colItems
' 		return1 = objItem.Hour   & ":" & objItem.Day  & ":" &  objItem.Month & ":" & objItem.Year
' 	 Next
' 	 'Exec local query
	
' 	Set colItems = objWMILocal.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
' 	For Each objItem in colItems
' 		return2 = objItem.Hour   & ":" & objItem.Day  & ":" &  objItem.Month & ":" & objItem.Year
' 	 Next
' 	 if return1 = return2 then
' 		return = "PASS"
' 	 Else
' 		return = "FAIL: HourDayMonthYear: " & return2
' 	 End If
' 	GetTimeMatch = return
' End Function
'Ping a remote system
' Function ReturnPing(strComputer)
' 	return = False
' 	On Error Resume Next
' 	Const wbemFlagReturnImmediately = &h10
' 	Const wbemFlagForwardOnly = &h20
' 	strQuery = "SELECT * FROM Win32_PingStatus Where Address = '" & strComputer & "'"
' 	Set objWMIService = GetObject("winmgmts:\\.\ROOT\CIMV2")
' 	Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
' 	For Each objItem in colItems
' 		if objItem.StatusCode = 0 Then
' 			return = True
' 		else
' 			return = False
' 		end if
' 	Next
' 	ReturnPing = return
' End Function
' 'Sets the IP Address of a local machine. (test)
' Function SetIP(strIP)
' 	strComputer = "."
' 	SearchPattern = "[B,Y,R,I][R,U,E,N][O,K,A,T][A,O,L,E]" ' For Broadcom, Yukon, and Realtek
' 	SQL = "Select MACAddress From Win32_NetworkAdapterConfiguration Where Caption LIKE '%" & SearchPattern & "%'"
' 	arrIPAddress = Array(strIP)
' 	arrSubnetMask = Array("255.255.0.0")
' 	arrGateway = Array("0.0.0.0")
' 	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
' 	Set colNetAdapters = objWMIService.ExecQuery(SQL)
' 	For Each objNetAdapter in colNetAdapters
' 		 errEnableStatic = objNetAdapter.EnableStatic(arrIPAddress, arrSubnetMask)
' 		 errGateways = objNetAdapter.SetGateways(arrGateway)
' 	Next
' End Function
' 'Get's mac address of local NIC
' Function GetMac()
' 	'Check to make sure there is an active adaptor
' 	SearchPattern = "[B,Y,R,I][R,U,E,N][O,K,A,T][A,O,L,E]" ' For Broadcom, Yukon, and Realtek
' 	SQL = "Select MACAddress From Win32_NetworkAdapterConfiguration Where Caption LIKE '%" & SearchPattern & "%'"
' 	Set objNetStat = GetObject("winmgmts://./root/cimv2").ExecQuery(SQL)
' 	With objNetStat
' 		If objNetStat.Count <> 0 Then
			
' 			For Each obj In objNetStat
' 				sAns = obj.MACAddress
' 			Next
' 			GetMac = sAns
' 		End If
' 	End With
' End Function

' 'Create a file with specific content
' Function ReplaceFile(strFileName,strContent)
' 	Const AppendIfExists = False
' 	Const ForReading = 1
' 	Const ForAppending = 8
' 	Const ForWriting = 2
' 	'on error resume next
' 	Set FS=CreateObject("Scripting.FileSystemObject")
' 	Set fFile = FS.OpenTextFile(strFileName,ForWriting,AppendIfExists)
' 	fFile.Write(strContent)
' 	fFile.Close
' End Function

Function MoveFile(strFileToCopy, strFolder)
	return = false
	Const OverwriteExisting = TRUE
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(strFolder) Then
		objFSO.CopyFile strFileToCopy, strFolder, OverwriteExisting  
	Else
		sendLog LogFile, "Target Folder does not exist.", "Copy Error"
	End If
	MoveFile = return
End Function

Function RenameFile(strFileOld, strFileNew)
    return = false
    Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
    If Fso.FileExists(strFileNew) Then
        Fso.DeleteFile strFileNew
    End If
    If Fso.FileExists(strFileOld) Then
        Fso.MoveFile strFileOld, strFileNew
    End If
    RenameFile = return
End Function

'Finds current Screen width
Function ScreenWidth()
	On Error Resume Next
	Const wbemFlagReturnImmediately = &h10
	Const wbemFlagForwardOnly = &h20
	strQuery = "SELECT * FROM Win32_DesktopMonitor Where Availability =3"
	Set objWMIService = GetObject("winmgmts:\\.\ROOT\CIMV2")
	Set colItems = objWMIService.ExecQuery(strQuery, "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
	For Each objItem in colItems
		return= objItem.ScreenWidth
	Next
	ScreenWidth = return
End Function