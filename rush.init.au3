#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <IE.au3>

Func _Au3RecordSetup()
Opt('WinWaitDelay',100)
Opt('WinDetectHiddenText',1)
Opt('MouseCoordMode',0)
Local $aResult = DllCall('User32.dll', 'int', 'GetKeyboardLayoutNameW', 'wstr', '')
If $aResult[1] <> '00000409' Then
  MsgBox(64, 'Warning', 'Recording has been done under a different Keyboard layout' & @CRLF & '(00000409->' & $aResult[1] & ')')
EndIf

EndFunc

Func _WinWaitActivate($title,$text,$timeout=10)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc

_AU3RecordSetup()
#endregion --- Internal functions Au3Recorder End ---

_WinWaitActivate("Program Manager","FolderView")
MouseClick("left",31,696,2)
MouseClick("left",29,484,2)

_WinWaitActivate("Microsoft Excel - BETA RUSHMORE AUTO-LOG","")
Send("{LWINDOWN}{RIGHT}{LWINUP}")

_WinWaitActivate("-Blue Sky Portal Login- - Windows Internet Explorer","Address Combo Contro")
Send("{LWINDOWN}{LEFT}{LWINUP}")

Send("tx1310t098{TAB}")
_WinWaitActivate("-Blue Sky Portal Login- - Windows Internet Explorer","Address Combo Contro")
Send("000000{ENTER}{ENTER}")

_WinWaitActivate("Portal (TX1310T098) - Windows Internet Explorer","Address Combo Contro")
MouseClick("left",377,64,1)
MouseClick("left",576,64,1)
MouseClick("left",579,120,1)
MouseClick("left",699,65,1)

_WinWaitActivate("Untitled Page - Windows Internet Explorer","")
Send("{SHIFTDOWN}tx{SHIFTUP}1310{SHIFTDOWN}t{SHIFTUP}098{TAB}")
MouseClick("left",335,41,1)

_WinWaitActivate("Portal (TX1310T098) - Windows Internet Explorer","Address Combo Contro")
MouseClick("left",576,64,1)
MouseClick("left",605,31,1)
MouseClick("left",699,65,1)
MouseClick("left",335,41,1)

_WinWaitActivate("Portal (TX1310T098) - Windows Internet Explorer","Address Combo Contro")
MouseClick("left",576,64,1)
MouseClick("left",605,75,1)
MouseClick("left",680,65,1)
MouseClick("left",335,250,1)
MouseClick("left",460,310,1)
MouseClick("left",460,325,1)
MouseClick("left",695,310,1)
MouseClick("left",695,395,1)
MouseClick("left",250,335,1)
MouseClick("left",690,845,1)
MouseClick("left",670,305,1)

#endregion --- Au3Recorder generated code End ---
