
'Global and Enviromental Variables'
    'message file used by the Select Case in refresh function'
        msgFile = GetCurDir() & "msg.txt"

    'Log File from update process. Needs to match running vbs script'
        logFile = "c:\temp\Log.txt"
        strBuild = ReadFile(GetCurDir() & "bin/bginfo/BGIdata/BuildVersion.txt")
        guiLog = "c:\temp\GuiLog.txt"
        
    'image line files
        imgFail ="bin/images/redx.png" 
	    imgPass ="bin/images/greencheck.png"
        imgGo ="bin/images/waitCover.gif"
        startImg = "bin/images/start.png"
    
'this function loads on first open of application	
    Sub Window_OnLoad
        'set visibility of form elements
            logSpan.style.visibility = "hidden"
            'EASpan.Style.visibility = "hidden"
            TMSpan.Style.visibility = "hidden"
            InitSpan.Style.visibility = "hidden"    
   
        'sets form properties depending on which device is running
        if Left(MBMFG(),4) = "Euro" Then
            systemType = "Handheld"
            logtext.style.fontSize = "15pt"
            logtext.cols = "75"
            logtext.rows = "8"
            defaultFolder = "c:\users\xAdministrator\Desktop"
           
        else
            systemType = "Client"
            logtext.rows = "10"
            logtext.cols = "120"
            defaultFolder = "c:\users\Administrator\Desktop"            
        End if       
        self.Focus()
    End Sub

