'========= PROCESS SUBS ==================
'this executes the select case from a button
Sub startSetup()
	CreateAFile msgFile,"INIT"
End Sub

'cleanup and shut down the system
Sub Success() 
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
End Sub