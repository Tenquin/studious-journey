
'List of Local State Message Variables
const lowbeam = "alpha(opacity=30)"
const hibeam = "alpha(opacity=100)"
intCounter = 0

'default msState
    msgState = "INIT"

'This is Main refresh timer Sub routine to keep the form moving along
Function MainLoop()
    
    'Select case which puts the GUI in the correct mode for each stage. Hide form elements, move them, react to conditions
    Select Case msgState
        'Initial Case at startup
        ' Case "INIT" 'setting up enviroment
        '    ' do things here to setup your applicaiton enviroment
        '     tSetMsg "Initiating System Please Wait..." 
        '     AppTitle.InnerHTML = "Current msgState = " & msgState
        '    'Move to another state
        '    intCounter = 5
        '     msgState = "START"            
            
        ' Case "START" 'Starting state of the app. Expects to see files and optical drive on with recovrey files on system.
        '    if intCounter <> 0 Then
        '         tSetMsg "Counter now = " & intCounter
        '         intCounter = intCounter - 1
        '     else
        '         tSetMsg "Final Counter reached"
        '         msgState = "END"
        '      end if
        '      AppTitle.InnerHTML = "Current msgState = " & msgState

        ' case "THIS"

    Case "STARTING"
        'Check if Build is already installed. Change If statement to validate
        Window_Onload              
        'this is first state if all is well
        ProcessList.InnerHTML = "EFEC system is ready to be upgraded."           
        'InitSpan.InnerHTML = "STARTING SYSTEM UPGRADE" 
        'TMSpan.style.filter = "alpha(opacity=60)"
        'InitSpan.style.visibility = show
        'Set all icons to default image
        Call SetIconDefaults(startImg)       
        ProcessList.style.color = "white"                  
        Call self.focus()
        
    Case "INIT"
        'make things happen in order to allow upgrade to procede. 
        'InitSpan.style.visibility = hide
        call RefreshLog(logFile)                     
        ProcessList.InnerHTML = "Setup process is getting ready to startup(Stopping Services)..."
        ProcessList.style.color = "white"

    Case "START_UTC_TIMEZONE" 'UTC_TIMEZONE CHECKPOINT 1
        taskTMimg1.src = imgGo
        taskTM1.style.color = "black"
        call RefreshLog(logFile)
        ProcessList.InnerHTML = "Started Setting UTC Timezone:"

    Case "FAIL_UTC_TIMEZONE"
        taskTMimg1.src = imgFail
        call RefreshLog(logFile)
        ProcessList.InnerHTML = "Failed Setting UTC Timezone:" 
        taskTM1.style.color = "red"           

    Case "PASS_UTC_TIMEZONE"
        taskTMimg1.src = imgPass
        call RefreshLog(logFile)
        ProcessList.InnerHTML = "Successfully Set UTC Timezone:"
        taskTM1.style.color = "lightGreen"  

    Case "START_OSF_REG" 'OSF_REG CHECKPOINT 2
        taskTMimg2.src = imgGo
        call RefreshLog(logFile)
        ProcessList.InnerHTML = "Started OSForensics Registration:" 
        taskTM2.style.color = "black"

    Case "FAIL_OSF_REG"
        taskTMimg2.src = imgFail
        call RefreshLog(logFile) 
        ProcessList.InnerHTML = "Failed OSForensics Registration:" 
        taskTM2.style.color = "red" 

    Case "PASS_OSF_REG" 
        taskTMimg2.src = imgPass
        call RefreshLog(logFile) 
        ProcessList.InnerHTML = "Passed OSForensics Registration:"
        taskTM2.style.color = "lightGreen"

    Case "START_SET_COMP_NAME"  'SET_COMP_NAME CHECKPOINT 3
        taskTMimg3.src = imgGo
        call RefreshLog(logFile) 
        ProcessList.InnerHTML = "Started Setting Computer Name:"  
        taskTM3.style.color = "black"

    Case "FAIL_SET_COMP_NAME"
        taskTMimg3.src = imgFail
        call RefreshLog(logFile) 
        ProcessList.InnerHTML = "Failed Setting Computer Name:" 
        taskTM3.style.color = "red" 

    Case "PASS_SET_COMP_NAME"
        taskTMimg3.src = imgPass
        call RefreshLog(logFile)
        ProcessList.InnerHTML = "Successfully Set Computer Name:"
        taskTM3.style.color = "lightGreen"

    Case "START_SET_BACKGROUND" 'SET_BACKGROUND CHECKPOINT 4
        taskTMimg4.src = imgGo
        call RefreshLog(logFile) 
        ProcessList.InnerHTML = "Started Setting Background Image:" 
        taskTM4.style.color = "black"

    Case "FAIL_SET_BACKGROUND"
        taskTMimg4.src = imgFail
        call RefreshLog(logFile)
        ProcessList.InnerHTML = "FAILED Setting Background Image:"
        taskTM4.style.color = "red" 

    Case "PASS_SET_BACKGROUND"
        taskTMimg4.src = imgPass
        call RefreshLog(logFile)
        ProcessList.InnerHTML = "Successfully Set Background Image:"
        taskTM4.style.color = "lightGreen"  

    Case "START_ACTIVATE_WINDOWS" 'ACTIVATE_WINDOWS CHECKPOINT 5
        taskTMimg5.src = imgGo
        call RefreshLog(logFile)         
        taskTM5.style.color = "black"
        ProcessList.InnerHTML = "Started Activating Windows:"  
        
    Case "FAIL_ACTIVATE_WINDOWS"
        taskTMimg5.src = imgFail
        call RefreshLog(logFile) 
        taskTM5.style.color = "red" 
        ProcessList.InnerHTML = "FAILED Activating Windows:"

    Case "PASS_ACTIVATE_WINDOWS"
        taskTMimg5.src = imgPass
        call RefreshLog(logFile)
        taskTM5.style.color = "lightGreen"
        ProcessList.InnerHTML = "Successfully Activated Windows:"

    Case "START_ACTIVATE_OFFICE" 'ACTIVATE_OFFICE CHECKPOINT 6
        taskTMimg6.src = imgGo
        call RefreshLog(logFile)
        taskTM6.style.color = "black" 
        ProcessList.InnerHTML = "Started Activating Office"

    Case "FAIL_ACTIVATE_OFFICE"
        taskTMimg6.src = imgFail
        call RefreshLog(logFile) 
        taskTM6.style.color = "red" 
        ProcessList.InnerHTML = "FAILED Activating Office"

    Case "PASS_ACTIVATE_OFFICE"
        taskTMimg6.src = imgPass
        call RefreshLog(logFile) 
            taskTM6.style.color = "lightGreen"  
            ProcessList.InnerHTML = "Successfully Activated Office"


    Case "START_IMPORT_REG_KEYS" 'IMPORT_REG_KEYS CHECKPOINT 7
        taskTMimg7.src = imgGo
        call RefreshLog(logFile)
        taskTM7.style.color = "black" 
        ProcessList.InnerHTML = "Started Importing Registry Keys:"

    Case "FAIL_IMPORT_REG_KEYS"
        taskTMimg7.src = imgFail
        call RefreshLog(logFile) 
        taskTM7.style.color = "red" 
        ProcessList.InnerHTML = "FAILED Importing Registry Keys:"

    Case "PASS_IMPORT_REG_KEYS"
        taskTMimg7.src = imgPass
        call RefreshLog(logFile)
            taskTM7.style.color = "lightGreen"  
        ProcessList.InnerHTML = "Successfully Importing Registry Keys:"

    Case "END"
        tSetMsg "END OF CASES. Click Left Gear to reset"
        AppTitle.InnerHTML = "Current msgState = " & msgState
        jShow_DivPowerPanel
    Case "EXIT"

    
        
        self.close
        
    Case Else
        
        tSetMsg "Main Loop failure: Msg:{" & msgState & "}"
                                
    End Select

End Function
