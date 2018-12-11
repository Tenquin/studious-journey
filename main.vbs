
const show = "visible"
const hide = "hidden"
const lowbeam = "alpha(opacity=30)"
const hibeam = "alpha(opacity=100)"

'This is Main refresh timer Sub routine to keep the form moving along
Sub RefreshList
    'reads the messege file to alter behavior
    if FileExists(msgFile) Then
        readMsg = ReadFile(msgFile)	
    Else
        'create MsgFile  in starting state if it doesn't exist
        CreateAFile msgFile,"STARTING"
        readMsg = ReadFile(msgFile)	
    End If

    'messaging debug log
     'SendLog guiLog,"ReadFile:" & readMsg,"READ"
    'ProcessList.InnerHTML = "Testing State" & readMsg 
    'Select case which puts the GUI in the correct mode for each stage of the update.
    Select Case readMsg
        Case "STARTING"
            'Check if Build is already installed. Change If statement to validate
            Window_Onload              
            'this is first state if all is well
            ProcessList.InnerHTML = "EFEC system is ready for ECP 007."           
            InitSpan.InnerHTML = "STARTING ECP 007" 
            'TMSpan.style.filter = "alpha(opacity=60)"
            InitSpan.style.visibility = show
            'Set all icons to default image
            Call SetIconDefaults(startImg)       
            ProcessList.style.color = "white"                  
            Call self.focus()
            
        Case "INIT"
            'make things happen in order to allow upgrade to procede. 
            InitSpan.style.visibility = hide
            call RefreshLog(logFile)
            TMSpan.style.visibility = show
            TMSpan.style.filter = lowbeam                    
            ProcessList.InnerHTML = "Setup process is getting ready to startup(Stopping Services)..."
            ProcessList.style.color = "white"
       
        Case "START_UTC_TIMEZONE" 'UTC CHECKPOINT 1
            TMSpan.style.filter = hibeam
            taskTMimg0.src = imgGo
            taskTM0.style.color = "black"
            call RefreshLog(logFile)
            ProcessList.InnerHTML = "Started Setting UTC Timezone:"

        Case "FAIL_UTC_TIMEZONE"
            TMSpan.style.filter = hibeam
            taskTMimg0.src = imgFail
            call RefreshLog(logFile)
            ProcessList.InnerHTML = "Failed Setting UTC Timezone:" 
            taskTM0.style.color = "red"           

        Case "PASS_UTC_TIMEZONE"
            TMSpan.style.filter = hibeam
            taskTMimg0.src = imgPass
            call RefreshLog(logFile)
            ProcessList.InnerHTML = "Successfully Set UTC Timezone:"
            taskTM0.style.color = "lightGreen"  

        Case "START_OSF_REG" 'OSF_REG CHECKPOINT 2
            taskTMimg1.src = imgGo
            call RefreshLog(logFile)
            ProcessList.InnerHTML = "Started OSForensics Registration:" 
            taskTM1.style.color = "black"

        Case "FAIL_OSF_REG"
            taskTMimg1.src = imgFail
            call RefreshLog(logFile) 
            ProcessList.InnerHTML = "Failed OSForensics Registration:" 
            taskTM1.style.color = "red" 

        Case "PASS_OSF_REG" 
            taskTMimg1.src = imgPass
            call RefreshLog(logFile) 
            ProcessList.InnerHTML = "Passed OSForensics Registration:"
            taskTM1.style.color = "lightGreen"

        Case "START_SET_COMP_NAME"  'SET_COMP_NAME CHECKPOINT 3
            taskTMimg2.src = imgGo
            call RefreshLog(logFile) 
            ProcessList.InnerHTML = "Started Setting Computer Name:"  
            taskTM2.style.color = "black"

        Case "FAIL_SET_COMP_NAME"
            taskTMimg2.src = imgFail
            call RefreshLog(logFile) 
            ProcessList.InnerHTML = "Failed Setting Computer Name:" 
            taskTM2.style.color = "red" 

        Case "PASS_SET_COMP_NAME"
            taskTMimg2.src = imgPass
            call RefreshLog(logFile)
            ProcessList.InnerHTML = "Successfully Set Computer Name:"
            taskTM2.style.color = "lightGreen"

        Case "START_SET_BACKGROUND" 'SET_BACKGROUND CHECKPOINT 4
            taskTMimg3.src = imgGo
            call RefreshLog(logFile) 
            ProcessList.InnerHTML = "Started Setting Background Image:" 
            taskTM3.style.color = "black"

        Case "FAIL_SET_BACKGROUND"
            taskTMimg3.src = imgFail
            call RefreshLog(logFile)
            ProcessList.InnerHTML = "FAILED Setting Background Image:"
            taskTM3.style.color = "red" 

        Case "PASS_SET_BACKGROUND"
            taskTMimg3.src = imgPass
            call RefreshLog(logFile)
            ProcessList.InnerHTML = "Successfully Set Background Image:"
            taskTM3.style.color = "lightGreen"  

        Case "START_ACTIVATE_WINDOWS" 'ACTIVATE_WINDOWS CHECKPOINT 5
            taskTMimg4.src = imgGo
            call RefreshLog(logFile)         
            taskTM4.style.color = "black"
            ProcessList.InnerHTML = "Started Activating Windows:"  
            
        Case "FAIL_ACTIVATE_WINDOWS"
            taskTMimg4.src = imgFail
            call RefreshLog(logFile) 
            taskTM4.style.color = "red" 
            ProcessList.InnerHTML = "FAILED Activating Windows:"

        Case "PASS_ACTIVATE_WINDOWS"
            taskTMimg4.src = imgPass
            call RefreshLog(logFile)
            taskTM4.style.color = "lightGreen"
            ProcessList.InnerHTML = "Successfully Activated Windows:"

        Case "START_ACTIVATE_OFFICE" 'ACTIVATE_OFFICE CHECKPOINT 6
            taskTMimg5.src = imgGo
            call RefreshLog(logFile)
            taskTM5.style.color = "black" 
            ProcessList.InnerHTML = "Started Activating Office"

        Case "FAIL_ACTIVATE_OFFICE"
            taskTMimg5.src = imgFail
            call RefreshLog(logFile) 
            taskTM5.style.color = "red" 
            ProcessList.InnerHTML = "FAILED Activating Office"

        Case "PASS_ACTIVATE_OFFICE"
            taskTMimg5.src = imgPass
            call RefreshLog(logFile) 
             taskTM5.style.color = "lightGreen"  
             ProcessList.InnerHTML = "Successfully Activated Office"


        Case "START_IMPORT_REG_KEYS" 'IMPORT_REG_KEYS CHECKPOINT 7
            taskTMimg6.src = imgGo
            call RefreshLog(logFile)
            taskTM6.style.color = "black" 
            ProcessList.InnerHTML = "Started Importing Registry Keys:"

        Case "FAIL_IMPORT_REG_KEYS"
           taskTMimg6.src = imgFail
            call RefreshLog(logFile) 
            taskTM6.style.color = "red" 
            ProcessList.InnerHTML = "FAILED Importing Registry Keys:"

        Case "PASS_IMPORT_REG_KEYS"
           taskTMimg6.src = imgPass
            call RefreshLog(logFile)
            taskTM6.style.color = "lightGreen"  
            ProcessList.InnerHTML = "Successfully Importing Registry Keys:" 
         
        Case "SUCCESS"
            TMSpan.style.filter = lowbeam
            InitSpan.InnerHTML = "<BR>SUCCESS<BR><BR>ECP 007 has completed."
            InitSpan.Style.BackgroundColor = "rgb(108, 161, 28)"
            InitSpan.Style.Color = "black"
            InitSpan.Style.visibility = show
            InitSpan.Style.height = "200px"
           
        Case "FAILURE"
            Call CopyFile(logFile,EnvPath("%userprofile%") & "\Desktop\FailureLog.txt")
            TMSpan.style.filter = lowbeam
            InitSpan.InnerHTML = "<BR>FAILURE<BR><BR>Contact Support Center with Log File"
            InitSpan.Style.BackgroundColor = "red"
            InitSpan.Style.Color = "white"
            InitSpan.Style.visibility = show
            InitSpan.Style.height = "200px"
            InitSpan.Style.top = "15%"
            MsgBox "An Error log file has been saved to your Desktop" & vblf & "Send this file to the Support Center for further assistances",0,"SETUP FAILURE"
            CreateAFile msgFile, "RESET"
            self.close            
        
        Case "RESET"
            Call DeleteFile(msgFile)
            Call DeleteFile(logFile)

        Case "EXIT"
            
            self.close

        Case Else
           
            ProcessList.InnerHTML = "Main Loop failure: Msg:{" & readMsg & "}"
            'redx.style.visibility = "visible"
            SendLog logFile,"MAIN LOOP FAILURE Msg:{" & readMsg & "}","MAIN_LOOP"
            refreshLog logFile          
                        
    End Select

End Sub
