'================ VARS ======================
'Product Code for Office required for activation
OffProdID = "XXXX-XXXX-XXXX-XXXX" 
'OffActID = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 'don't think this is needed
OfficeTXT = "c:\OfficeIDs\confid.txt"
Serial = MBSerial()
confID = GetConfID()
'ori Matching pattern
'==================== EXECUTE ================
If Not WScript.Arguments.Named.Exists("elevate") Then
	CreateObject("Shell.Application").ShellExecute WScript.FullName _
	  , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
	WScript.Quit
End If
'Activate License if neccessary
isOfficeLicensed()
'Get Current Office ACTIVATION status
curActivation = isOfficeAct()
'Make sure office isn't already activated then activate office		
if curActivation = False Then
	
	activateOff = ActivateOffice(confID)
		
		'validate office activation after attempted activation
		If activateOff = "SUCCESS" Then
			Wscript.Echo "Successfully activated Office"
		Else
			errorFlag = True
			Wscript.Echo "FAILED Office Activation with stored ConfID:" & activateOff
		End If
Else
	'Office already Activated
	Wscript.Echo "Office has already been Activated."
End If

'MsgBox isOfficeAct()
'MsgBox GetOfficeID()
'================ FUNCTIONS ==================
'Activate office with a ConfirmationID
Function ActivateOffice(strConfID)
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
	tmpResult = GetBINResult(binFile & strConfID)
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

Function GetConfID()
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(OfficeTXT,1)
	strResult = objFileToRead.ReadAll()
	strPattern = "(\d{48})"
	GetConfID = REGEX_GetMatch(strResult, strPattern)
	objFileToRead.Close
	Set objFileToRead = Nothing
End Function

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
	path="C:\temp\"   
	exists = fso.FolderExists(path)
	if NOT (exists) then 
		fso.CreateFolder "C:\temp"
	End If
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
		objShell.run("cscript " & chr(34) & "c:\Program Files (x86)\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /inpkey:" & OffProdID)
	End If
	if (FileExists(file64)) then
		objShell.run("cscript " & chr(34) & "c:\Program Files\Microsoft Office\" & newfolder & "\ospp.vbs" & chr(34) & " /inpkey:" & OffProdID)
	End If
End Function
