#include <GUIConstantsEx.au3>
#include <MsgBoxConstants.au3>
#include <IE.au3>

romaintain()


Func romaintain()															; GUI

	Local $hGUI = GUICreate("AUTOTOOL - BIN LOC.", 250, 100)				; Create a GUI with various controls.

	Local $itechID, $iSNUM, $ibenchLOC    									; Initialize Vars : text input

	Local $techID, $SNUM, $benchLOC											; Initialize Vars : pass to roupdate

    GUICtrlCreateLabel("Serial Number:", 10, 20)							; Label and Entry fields
	$isNUM = GUICtrlCreateInput("", 100, 15, 125, 20)

	GUICtrlCreateLabel("Tech ID :", 10, 50)
	$itechID = GUICtrlCreateInput("", 100, 45, 125, 20)

	GUICtrlCreateLabel("Bench Location:", 10, 80)
	$ibenchLOC = GUICtrlCreateInput("", 100, 75, 50, 20)

	Local $iUPDATE = GUICtrlCreateButton( "Update RO", 160, 75, 75, 25)		; Create Update button

	GUISetState(@SW_SHOW, $hGUI)											; Display the GUI.

	While 1																	; Loop until the user exits.
		Switch GUIGetMsg()

			Case $iUPDATE

			   $benchLOC = GUICtrlRead( $ibenchLOC )						; Copy user inputs to Var
			   $techID = GUICtrlRead( $itechID )
			   $sNUM = GUICtrlRead( $isNUM )

			   roupdate("K1R5" ,$sNUM , "WU2" )  							; Call roupdate

			   Sleep(1000)													; Wait for Unique Dialog to load

			   roassign($benchLOC, $techID) 								; Call roassign

			   Sleep(1000)
			   MsgBox(64, "Rushmore RO Maintainer", "SUCCESS! Bin location changed to " & $benchLOC & ", and assigned to " & $techID )

			Case $GUI_EVENT_CLOSE

			; Delete the previous GUI and all controls.
			GUIDelete($hGUI)

			ExitLoop
		EndSwitch
	 WEnd
 EndFunc

Func roupdate($iWH, $iSN, $iRoType) 										; Interfaces with Repair Order Maintain screen, pulls up Unique Dialog
   Local $oIE , $hwnd, $oForm, $oQuery, $oHref								; Initialize Local Variables

   $oIE = _IEAttach("User Interface")										; Attach to Repair Order Maintain page
   $hwnd = _IEPropertyGet($oIE, "hwnd")
   $oForm = _IEGetObjByName($oIE, "form1")

   $oQuery = _IEFormElementGetObjByName($oForm, "drpWH")					; Select WH
   _IEFormElementSetValue($oQuery, $iWH)

   $oQuery = _IEFormElementGetObjByName($oForm, "txtSN")					; Enter SN
   _IEFormElementSetValue($oQuery, $iSN)

   $oQuery = _IEFormElementGetObjByName($oForm, "drpRoType")				; Select RO Type
   _IEFormElementSetValue($oQuery, $iRoType)

   $oQuery = _IEFormElementGetObjByName($oForm, "btnSearch")				; Commence search for unit
   _IEAction($oQuery, "Click")

   Sleep(1000)																; Allow dynamic table to load

   $oQuery = _IEGetObjById($oForm, "gvROLine")								; Change focus to table, get unique HREF
   $oHref = _IEGetObjById($oQuery, "gvROLine_ctl02_lnkEnger")

   _IEAction($oHref, "focus")												; Select and click the tech ID,
   ControlSend($hwnd, "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "{Enter}")

EndFunc

Func roassign($iBinLoc, $iAsgEnger)											; Update with user Bench # and Tech ID
Local $oIE, $hwnd, $oForm, $oQuery, $sBin

$oIE = _IEAttach("Untitled Page", "DialogBox")
$oForm = _IEGetObjByName($oIE, "form1")

$oQuery = _IEFormElementGetObjByName($oForm, "txtReBin")
$sBin = _IEFormElementGetValue ($oQuery)
$sBin = StringMid( $sBin, 1, 8 )



_IEFormElementSetValue($oQuery, $sBin & "-" & $iBinLoc )

$oQuery = _IEFormElementGetObjByName($oForm, "drpAsgEnger")
_IEFormElementSetValue($oQuery, $iAsgEnger)

$oQuery = _IEFormElementGetObjByName($oForm, "btnSave")
_IEAction($oQuery, "Click")

Sleep (1000)
WinClose("[CLASS:Internet Explorer_TridentDlgFrame]")


EndFunc

