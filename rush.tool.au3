#include <GUIConstantsEx.au3>

#include <MsgBoxConstants.au3>
#include <IE.au3>

rshmrtool ()


Func rshmrtool ()

   Local $hGUI = GUICreate("Rushmore Auto Tool", 400, 180)	   ; Create a GUI with various controls.

   GUICtrlCreateLabel("---Login to MyService---", 22, 2)

   	Local $itechID, $iPW, $isNUM, $ipNUM, $ibenchLOC						; Initialize Vars : text input

	Local $techID, $PW, $rushCM, $rushCS, $sNUM, $pNUM, $benchLOC			; Initialize Vars : pass to roupdate

    GUICtrlCreateLabel("Tech ID :", 10, 25)									; Label and Entry fields
	$itechID = GUICtrlCreateInput("", 70, 20, 75, 20)

	GUICtrlCreateLabel("Password :", 10, 45)
	$iPW = GUICtrlCreateInput("", 70, 42, 75, 20)

    Local $iLOGIN = GUICtrlCreateButton( "LOGIN", 4, 65, 75, 25)			; Create Login button
	Local $iLOGOFF = GUICtrlCreateButton( "LOGOFF", 80, 65, 75, 25)			; Create Logoff button

	GUICtrlCreateLabel("---Check Output Report---", 22, 95)
	$rushCM = GUICtrlCreateRadio("CM", 80, 110, 35, 20)
    $rushCS = GUICtrlCreateRadio("CS", 40, 110, 30, 20)
    Local $iOUTPUT = GUICtrlCreateButton( "OUTPUT REPORT", 20, 130, 115, 25); Create Update button
	GUICtrlSetState($rushCM, $GUI_CHECKED)

    GUICtrlCreateLabel("---version .01 by JR---",25, 160)

	GUICtrlCreateLabel("Serial Number:", 170, 7)
	$isNUM = GUICtrlCreateInput("", 250, 5, 125, 20)

    Local $iSYSREP = GUICtrlCreateButton( "SYSTEM REPAIR", 160, 28, 100, 25)
    Local $iBOM = GUICtrlCreateButton( "BOM", 260, 28, 40, 25)
    Local $iWURINQ = GUICtrlCreateButton( "CHECK STATUS", 300, 28, 100, 25)

    GUICtrlCreateLabel("---Update Location in MyService---", 200, 55)
    GUICtrlCreateLabel("Bench Number:", 170, 75)
	$ibenchLOC = GUICtrlCreateInput("", 250, 72, 50, 20)

	Local $iUPDATE = GUICtrlCreateButton( "UPDATE", 310, 70, 75, 25)		; Create Update button

	GUICtrlCreateLabel("---Order Part---", 240, 97)
	GUICtrlCreateLabel("Part Number: ", 170, 116)
	$ipNUM = GUICtrlCreateInput("", 238, 113, 70, 20)
    Local $iORDER = GUICtrlCreateButton( "ORDER PART", 310, 111, 85, 25)	; Create Order Part button
    Local $iRECIEVE = GUICtrlCreateButton( "CONFIRM RECIEVE PART", 200, 140, 150, 25)
    ;GUICtrlCreateLabel("---Open Selected Page---", 225, 140)
	;Local $iComboBox = GUICtrlCreateCombo("Select a page", 170, 155, 140, 20)
	;Local $iOPEN = GUICtrlCreateButton( "OPEN", 310, 155, 75, 25)			; Create Open button


	GUISetState(@SW_SHOW, $hGUI)											; Display the GUI.

	While 1																	; Loop until the user exits.
		Switch GUIGetMsg()
			Case $iLOGIN
			   $techID = GUICtrlRead( $itechID )							; Copy user inputs to Var
			   $PW = GUICtrlRead( $iPW )
			   rshmrlogin ($techID, $PW) 									; Call rologin
			   Sleep(1000)													; Wait for 1000 ms
			   ;rshmrsetup ($techID)											; Setup MS screen

			Case $iLOGOFF
			      $oIE = _IEAttach("Portal")
				  _IEQuit($oIE)

			Case $iOUTPUT
			Case $iSYSREP
			Case $iBOM
			Case $iWURINQ
			Case $iUPDATE
			   $benchLOC = GUICtrlRead( $ibenchLOC )						; Copy user inputs to Var
			   $techID = GUICtrlRead( $itechID )
			   $sNUM = GUICtrlRead( $isNUM )
			   roupdate("K1R5" ,$sNUM , "WU2" )  							; Call roupdate
			   Sleep(1000)													; Wait for Unique Dialog to load
			   roassign($benchLOC, $techID) 								; Call roassign
			   Sleep(1000)
			   MsgBox(64, "Rushmore Auto Tool", "SUCCESS! Bin location changed to " & $benchLOC & ", and assigned to " & $techID )

			;Case $iOPEN
			Case $iORDER
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

Func rshmrlogin($iUID, $iPW)
   Local $oIE , $hwnd, $oForm, $oQuery, $oHref

   $oIE = _IECreate("http://myservice.servms.com/myService/WebLogin.aspx?LoginType=SMS", 0, 1, 1, 1)

   _IELoadWait($oIE)

   $oForm = _IEGetObjByName($oIE, "Form2")

   $oQuery = _IEFormElementGetObjByName($oForm, "txtUserID")
   _IEFormElementSetValue($oQuery, $iUID)
   $oQuery = _IEFormElementGetObjByName($oForm, "txtPassword")
   _IEFormElementSetValue($oQuery, $iPW)

   $oQuery = _IEFormElementGetObjByName($oForm, "cmdLogin")
   _IEAction($oQuery, "Click")

   Sleep(1000)

   ControlSend("", "", "[CLASS:Internet Explorer_Server; INSTANCE:1]", "{Enter}")
EndFunc

Func rshmrsetup($iUID)
   Local $oIE , $hwnd, $oForm, $oQuery

   _IECreate("http://myservice.servms.com/MYSERVICE/ER/ST/ERST001.ASPX", 1, 1, 1, 1) ;System Repair
   _IECreate("http://myservice.servms.com/MYSERVICE/ER/RO/ERRO001.ASPX", 1, 1, 1, 0) ;Repair Order create and maintain
   _IECreate("http://myservice.servms.com/MYSERVICE/ER/REPORT/ER003W0.ASPX", 1, 1, 1, 0) ;Engineer output report
   ;_IECreate("http://myservice.servms.com/MYSERVICE/SM/REPORT/SM016C0.ASPX?SERVICEM=P", 1, 1, 1, 0) ;WUR Status Inquiry

   $oIE = _IEAttach("Untitled")
   $oForm = _IEGetObjByName($oIE, "form1")

    $oQuery = _IEFormElementGetObjByName($oForm, "txtEngID")
   _IEFormElementSetValue($oQuery, $iUID)
EndFunc

Func rshmroutput($idpt)


EndFunc