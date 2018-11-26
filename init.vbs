'Global Variables'
    'app settings
    
    msgFile = GetCurDir() & "msg.txt"
    'Target OS Path. What will be captured
        curAppPath = GetCurDir()

        'logFile = "c:\temp\ECP17Log.txt"
        'strBuild = ReadFile(GetCurDir() & "bin/bginfo/BGIdata/BuildVersion.txt")
        'guiLog = "c:\temp\GuiLog.txt"
     'Log File from update process. Needs to match running vbs script'
        logFile = GetCurDir() & "Log.txt"
               
    'image line files
        imgFail ="bin/img/redx.png"
	    imgPass ="bin/img/greencheck.png"
        imgGo ="bin/img/waitCover.gif"
        startImg = "bin/img/start.png"
     
     'Kick off refresh timer on main.vbs loop
    intRefresh = 300
     iTimerID = window.setInterval("MainLoop()", intRefresh)
'this function loads on first open of application. 'does not function if using Jquery. 	
    Sub Window_OnLoad                             
        'self.Focus()
        'document.body.style.backgroundImage = "url('img/efec.jpg')" 
        
    End Sub

