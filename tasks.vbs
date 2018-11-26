''========= PROCESS SUBS ==================
'this executes the select case from a button
Sub startSetup()
CreateAFile msgFile,"INIT"
End Sub

'cleanup and shut down the system
Sub tSuccess() 
' Perform any cleanup here

End Sub

'Save log file to desktop
Sub saveLog()

End Sub


'set all default task states . pass in image source
Sub SetIconDefaults(strImagePath)       
	'TM tasks
	taskTMimg1.src = strImagePath
	taskTMimg2.src = strImagePath
	taskTMimg4.src = strImagePath
	taskTMimg3.src = strImagePath
	taskTMimg4.src = strImagePath
	taskTMimg5.src = strImagePath
	taskTMimg6.src = strImagePath
	taskTMimg7.src = strImagePath
	taskTMimg8.src = strImagePath

	'EA tasks
	taskEAimg1.src = strImagePath
	taskEAimg2.src = strImagePath
	taskEAimg4.src = strImagePath
	taskEAimg3.src = strImagePath
	taskEAimg4.src = strImagePath
	taskEAimg5.src = strImagePath
	taskEAimg6.src = strImagePath
End Sub

'Set the source of a single image object
Sub SetIconImage(strObjectName, strImagePath)
strObjectName.src = strImagePath
End Sub

'refresh log on form
Sub RefreshLog(strFile)
 logtext.Value = ParseLog(ReadFile(strFile))
 logtext.scrollTop = logtext.scrollHeight
 logSpan.style.visibility = "visible"
 logtext.InnerHTML = "BLAH"
End Sub

'Change text in ProcessList HTML object
Function tsetMsg(strMessageText)
	ProcessList.InnerHTML = strMessageText
End Function