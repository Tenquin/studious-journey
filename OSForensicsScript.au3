;Developer: jchoover@scires.com, Scientific Research Corporation

#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#RequireAdmin
#include <WinAPIFiles.au3>


$InstallPath = "C:\Program Files\OSForensics\OSForensics.exe"
$WindowTitle = "PassMark® OSForensics"
$WindowRegister = "Register OSForensics"
$WindowOffer = "OSForensics Upgrade offer"
$WindowThanks = "Thanks"
$WindowForensics = "OSForensics"
;$FileName = "\\vmware-host\Shared Folders\VMWare Shared\OSForensicsKeyEntry.ini"
$FileName = @ScriptDir & "\OSForensicsKeyEntry.ini"
Local $var = WinList()

;run OSForensics
If FileExists($InstallPath)Then
   run($InstallPath,"",@SW_HIDE)
   ;RunAs("MCEDS Admin","domain.Local","oiuy9876LKJH)(*&", 0, $InstallPath)
Else
   MsgBox(4096, $InstallPath, "File: " & $InstallPath & " does NOT exist.")
   Exit
EndIf

;Local $sRead = StringReplace(IniRead($FileName, "OSF Key", "Key", "Default Value"), "{ENTER}", @CRLF))
Local $sRead1 = IniRead($FileName, "OSF Key", "Key_Line1", "Default Value")
Local $sRead2 = IniRead($FileName, "OSF Key", "Key_Line2", "Default Value")
Local $sRead3 = IniRead($FileName, "OSF Key", "Key_Line3", "Default Value")
Local $sRead4 = IniRead($FileName, "OSF Key", "Key_Line4", "Default Value")

WinActivate($WindowTitle, "Continue Using Trial Version")
Sleep(1000)
WinWaitActive($WindowTitle, "")
;If IsAdmin() Then MsgBox($MB_SYSTEMMODAL, "", "The script is running with admin rights.")
Sleep(1000)

;ControlClick($WindowTitle, "", "[ID:1036]")

If WinActive($WindowTitle)Then
   ControlClick($WindowTitle, "Upgrade to Professional Version", "Button2")

   If WinActive($WindowRegister)Then
	  ;ControlSend($WindowRegister, "", "[CLASS:Button; INSTANCE:5]", "This is some text")
	  Send($sRead1)
	  Send("{ENTER}")
	  Send("{#}")
	  Send($sRead2)
	  Send("{ENTER}")
	  Send($sRead3)
	  Send("{ENTER}")
	  Send($sRead4)
	  Send("{TAB}")
	  Send("{ENTER}")
	  If WinActive ($WindowOffer) Then
		 ;Sleep(1000)
		 ;Send("{TAB}")
		 Send("{ENTER}")
	  EndIf
	  If WinActive ($WindowThanks) Then
		 WinKill($WindowThanks)
		 WinKill($WindowTitle)
		 Sleep(1000)
		 WinKill($WindowForensics)
	  EndIf
   Else
	  MsgBox($MB_SYSTEMMODAL, "", "Registration Window Was Closed Before Completion.  Please exit OSForensics and restart the script.")
	  Exit
   EndIf
   ;MsgBox($MB_SYSTEMMODAL, "", "Sent Tab Enter.")
   ;Exit
Else
   MsgBox($MB_SYSTEMMODAL, "", "Registration Window Was Closed Before Completion.  Please restart the script.")
   Exit
EndIf

