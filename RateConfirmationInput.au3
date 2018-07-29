#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <String.au3>
#include <Constants.au3>
#include <GUIConstantsEx.au3>

AutoItSetOption("MouseCoordMode", 0)

;Mapping functions to hotkeys.
HotKeySet("{ESC}", "Terminate")
HotKeySet("^`", "BuildRateCon")
HotKeySet("^1", "EmailRate")
HotKeySet("^2", "CollectInfo")
HotKeySet("^3", "FillInfo")

;Delays, adjust to balance speed and error avoidance
Local $tabDelay = 50
Local $initialDelay = 5000
Local $newActiveLoadDelay = 5000
Local $locDelay = 2000
Global $emailSendDelay = 7000

;Variables
Global $WONumber
Global $rate
Global $shipper
Global $shipperDate
Global $shipperTimeDwnCnt
Global $comments
Global $consignee
Global $consigneeDate
Global $consigneeTimeDwnCnt
Global $poNumbers

;Formula values
Global $distance
Global $ratePerMile
Global $carrierFee
Global $profit

;Tonu boolean, takes different route.
Global $isTONU

;Collects info from Load Report
Func CollectInfo()

   $isTONU = False ;resets isTONU flag

   ;Opens and checks for Load Report
   If WinActivate("Load Report") = 0 Then
	  MsgBox(0, "Window Error", "Unable to access Load Report")
	  Return
   EndIf


   ;-----------------LEAN ID(WONumber)-----------------
   Send("{HOME}")
   Sleep(100)
   Send("/")
   Send(" - ")
   Send("^+{LEFT}")
   Sleep(50)
   Send("^c")
   Sleep(50)
   Send("{tab}")

   Local $WONumErr = 0

   $WONumber = ClipGet()

   If $WONumber < 60000000 Then

	  $WONumErr = MsgBox(4, "Warning", "WO number: " & $WONumber & "might be incorrect. Continue?")

   EndIf

   If $WONumErr = 7 Then

	  Exit

   EndIf


   ;-----------------RATE-----------------
   Send("/")
   Send("USD")
   Send("^+{UP}")
   Send("^c")
   Send("{tab}")

   $rate = StringStripWS(ClipGet(), 8)

   Local $RateErr = 0

   ;If rate is equal to WO number
   If $rate = $WONumber Then

	  $RateErr = MsgBox(6, "Warning", "Rate: " & $rate & " might be incorrect. Press Continue to build as TONU.")

   EndIf

   If $RateErr = 2 Then

	  Exit

   EndIf

   If $RateErr = 10 Then

	  BuildRateCon()
	  Return

   EndIf

   If $RateErr = 11 Then

	  $isTONU = True
	  $rate = 150

   EndIf

   ;-----------------SHIPPER-----------------
   Send("/")
   Send(" Pick ")
   Send("+{DOWN 3}")
   Send("^c")
   Send("{tab}")

   Local $pickLocArray = _StringExplode(ClipGet(), @CRLF, 0)

   ClipPut($pickLocArray[2]); copy clipboard value into array.

   $shipper = ClipGet()

   ;-----------------SHIPPER DATE-----------------
   Send("/")
   Send("PICK PLAN")
   Send("+{DOWN 2}")
   Send("^c")
   Send("{tab}")

   ;Parses date
   Local $dateArray2 = _StringExplode(StringMid(ClipGet(), 11), "/", 0)

   $shipperDate = "2018-" & $dateArray2[0] & "-" & $dateArray2[1]

   ;-----------------SHIPPER TIME-----------------
   Send("/")
   Send("APPT")
   Send("+{UP 3}")
   Send("^c")
   Send("{tab}")

   Local $TimeSplit = ClipGet()
   Local $Array1 = _StringExplode($TimeSplit, ":", 0)
   $shipperTimeDwnCnt = ($Array1[0] * 4) + ($Array1[1] / 15)


   ;-----------------COMMENTS-----------------
   Send("{HOME}")
   Sleep(100)
   Send("/")
   Send("Comments")
   Send("+{DOWN 2}")
   Send("^c")
   Send("{tab}")

   Local $aComments = _StringExplode(ClipGet(), @CRLF, 0)

   If StringInStr($aComments[1], "Stops") Or StringInStr($aComments[1], "Rate") Then

	  $comments = StringStripWS( StringMid($aComments[0], 9) , 1) ;Strips lead white space and takes string after "Comments"

   Else

	  $comments = StringStripWS( (StringMid($aComments[0], 9) & $aComments[1]) , 1)

   EndIf

   If $isTONU = True Then
	  $poNumbers = ""
   Else
   ;-----------------PO Numbers-----------------
	  ;------------SHIPPER REF #------------
	  Send("/")
	  Send("Shipper ref")
	  Send("+{DOWN}")
	  Send("+{LEFT}")
	  Send("^c")
	  Send("{tab}")

	  Local $shipRefNumber = StringMid(ClipGet(), 15)


	  ;------------SHIPMENT #------------
	  Send("/")
	  Send("Shipments: ")
	  Send("+{DOWN}")
	  Send("^c")
	  Send("{tab}")

	  Local $shipNumber = StringMid(ClipGet(), 12)

	  $poNumbers = $shipRefNumber & " / " & $shipNumber

   EndIf


   ;-----------------CONSIGNEE-----------------
   Send("/")
   Send(" Drop ")
   Send("+{DOWN 3}")
   Send("^c")
   Send("{tab}")

   Local $dropLocArray = _StringExplode(ClipGet(), @CRLF, 0)

   ClipPut($dropLocArray[2])

   $consignee = ClipGet()


   ;-----------------CONSIGNEE DATE-----------------
   Send("/")
   Send("DROP PLAN")
   Send("+{DOWN 2}")
   Send("^c")
   Send("{tab}")

   ;Takes the first 11 chars and parses them.
   Local $dateArray1 = _StringExplode(StringMid(ClipGet(), 11), "/", 0)

   $consigneeDate = "2018-" & $dateArray1[0] & "-" & $dateArray1[1]

   ;-----------------CONSIGNEE TIME-----------------
   Send("/")
   Send("APPT")
   Send("{f3}");goes to second result
   Send("+{UP 3}")
   Send("^c")
   Send("{tab}")

   Local $TimeSplit = ClipGet()
   Local $Array1 = _StringExplode($TimeSplit, ":", 0)
   $consigneeTimeDwnCnt = ($Array1[0] * 4) + ($Array1[1] / 15)

   ;-----------------PROFIT ALG-----------------
   If _IsChecked($profitCheckBox) Or GUICtrlRead($idInput) > 0 Then
	  ProfitAlgorithm()
   EndIf

EndFunc

;To be run in collectInfo Func
Func ProfitAlgorithm()

   ;-----------------DISTANCE-----------------
   Send("/")
   Send(" mi ")
   Send("+{UP}")
   Send("^c")
   Send("{tab}")

   Local $aRate

   If StringInStr($rate, ",") Then

	  $aRate = _StringExplode($rate, ",", 0)
	  $rate = $aRate[0]*1000 + $aRate[1]

   EndIf

   $distance = StringStripWS(ClipGet(), 8)
   $ratePerMile = (Int($rate) / $distance)

   ConsoleWrite("Distance: " & $distance)
   ConsoleWrite(" Rate: " & $rate)
   ConsoleWrite(" RPM: " & $ratePerMile)


   If $ratePerMile >= 2.5 And $ratePerMile < 3 Then

	  $carrierFee = Round($rate * .99, 2)

   ElseIf $ratePerMile >= 3 And $ratePerMile < 3.5 Then

	  $carrierFee = Round($rate * .98, 2)

   ElseIf $ratePerMile >= 3.5 And $ratePerMile < 4 Then

	  $carrierFee = Round($rate * .97, 2)

   ElseIf $ratePerMile >= 4 And $ratePerMile < 4.5 Then

	  $carrierFee = Round($rate * .96, 2)

   ElseIf $ratePerMile >= 4.5 And $ratePerMile < 5 Then

	  $carrierFee = Round($rate * .95, 2)

   ElseIf $ratePerMile >= 5 And $ratePerMile < 5.5 Then

	  $carrierFee = Round($rate * .93, 2)

   ElseIf $ratePerMile >= 5.5 Then

	  $carrierFee = Round($rate * .90, 2)

   EndIf


EndFunc

;Note: values have been modified to protect confidentiality

Func FillInfo()

   WinActivate("ITS DISPATCH")

   ;-----------------New Active Load-----------------
   Send("{TAB 26}")

   Send("{enter}")

   Sleep($newActiveLoadDelay)


   ;-----------------Bottom PO Section-----------------
   Send("+{TAB 34}")

   Send($poNumbers)


   ;-----------------CONSIGNEE TIME-----------------
   Send("+{TAB 7}")

   For $i = 1 To $consigneeTimeDwnCnt Step +1

	  Send("{DOWN}")

   Next


   ;-----------------CONSIGNEE DATE-----------------
   Send("+{TAB}")
   ClipPut($consigneeDate)
   Send("{APPSKEY}")
   Sleep(50)
   Send("{DOWN 3}")
   Send("{ENTER}")


   ;-----------------CONSIGNEE-----------------
   Send("+{TAB 3}")
   Send($consignee)
   Sleep($locDelay)


   ;-----------------Top PO section-----------------
   Send("+{TAB 6}")
   Send($poNumbers)


   ;-----------------COMMENTS-----------------
   Send("+{TAB}")
   Send($comments)


   ;-----------------WEIGHT-----------------
   Send("+{TAB 3}")

   Send("42000")


   ;-----------------SHIPPER TIME-----------------
   Send("+{TAB 5}")

   For $i = 1 To $shipperTimeDwnCnt Step +1

	  Send("{DOWN}")

   Next


   ;-----------------CONSIGNEE DATE-----------------
   Send("+{TAB}")
   ClipPut($shipperDate)
   Send("{APPSKEY}")
   Sleep(50)
   Send("{DOWN 3}")
   Send("{ENTER}")

   ;-----------------SHIPPER-----------------
   Send("+{TAB 3}")
   Send($shipper)
   Sleep($locDelay)

   ;-----------------EQUIPMENT TYPE-----------------
   Send("+{TAB 7}")

   Send("{DOWN 4}")

   ;-----------------CARRIER-----------------
   Send("+{TAB}")

   Send("dm world trans")

   Sleep(1500)


   ;-----------------RATE-----------------
   Send("+{TAB 9}")
   Send($rate)


   ;-----------------TYPE for TONUs-----------------
   Send("+{TAB}")
   If $isTONU = True Then
	  Send("{DOWN 15}")
   EndIf

   ;-----------------LEAN ID-----------------
   Send("+{TAB 3}")
   If $isTONU = True Then
	  Send($WONumber & " TONU")
   Else
	  Send($WONumber)
   EndIf
   ;-----------------BILL TO-----------------
   Send("+{TAB 4}")
   Send("amazon")

   Sleep(1500)

   ;-----------------CARRIER FEE-----------------
   Send("{TAB 65}")
   If $isTONU = True Then
	  Send("{TAB}") ;an extra tab is needed for TONUs
   EndIf
   If _IsChecked($profitCheckBox) Then
	  ClipPut($carrierFee)
   ElseIf GUICtrlRead($idInput) > 0 Then
	  ClipPut($rate - ((GUICtrlRead($idInput) / 100) *$rate))
   Else
	  ClipPut($rate)
   EndIf

   Send("^v")
   Send("{ENTER}")

EndFunc

Func BuildRateCon()

   ;Opens and checks for ITS Dispatch
   If WinActivate("ITS DISPATCH") = 0 Then
	  MsgBox(0, "Window Error", "Unable to access ITS DISPATCH")
	  Return

   EndIf

   Send("{f5}")

   ;Turns caps lock off.
   Opt("SendCapslockMode", 0)
   Send("{CapsLock off}")

   CollectInfo()

   Sleep($initialDelay)

   FillInfo()

EndFunc

;-----------------Terminate Function-----------------
;To terminate when esc is pressed
Func Terminate()

   Exit

EndFunc

;-----------------Email Function-----------------
Func EmailRate()

   WinActivate("ITS DISPATCH")

   ;Get WO number + TONU label
   MouseClick($MOUSE_CLICK_LEFT, 1350, 375, 3, 1)
   Send("^c")

   ;Press Email Button
   MouseClick($MOUSE_CLICK_LEFT, 1090, 830, 1, 1)

   Sleep(1500)

   MouseClick($MOUSE_CLICK_LEFT, 840, 445, 1, 1)

   ;Email for dmworld@dmwtrans.com checkbox.
   If _IsChecked($dmCheckBox) Then

	  Send("example@example.com")

   ElseIf _IsChecked($meCheckBox) Then

	  Send("me@example.com")

   Else

	  Send(GUICtrlRead($emailAddr))

   EndIf


   Send("{tab}")
   Send("{tab}")

   Send("^v")


   If _IsChecked($msgCheckBox) Then

	  Send(" Past Loads Rate Confirmation for Data Entry Team")

   Else

	  Send(GUICtrlRead($emailMsg))

   EndIf

   If _IsChecked($autoSendCheckBox) Then

	  Send("{TAB 3}")
	  Send("{ENTER}")
	  Sleep($emailSendDelay)
	  Send("+{TAB 5}")
	  Send("{ENTER}")

   EndIf

EndFunc

;AutoIt written Function
Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc


;----------------------------------GUI----------------------------------
Local $hMainGUI = GUICreate("RCBot", 200, 230)

GUICtrlCreateLabel("Press 'ESC' at any time to end program.", 3, 5)

Local $buildRateConButton = GUICtrlCreateButton("Build Rate Con ( Ctrl + ~ )", 10, 35, 135, 25)

;Creates UpDown Selector
GUICtrlCreateLabel("%", 158, 23)
Global $idInput = GUICtrlCreateInput("0", 150, 36, 35, 22)
GUICtrlCreateUpdown($idInput)

Global $profitCheckBox = GUICtrlCreateCheckbox("Profit Formula", 10, 60)

Local $emailTemplateButton = GUICtrlCreateButton("Email Rate Con ( Ctrl + 1 )", 10, 85, 135, 25)

GUICtrlCreateLabel("Email Address", 10, 115)
Global $dmCheckBox = GUICtrlCreateCheckbox("Ex", 87, 114)
Global $meCheckBox = GUICtrlCreateCheckbox("Me", 125, 114)
Global $emailAddr = GUICtrlCreateInput("", 10, 135, 150, 20)

GUICtrlCreateLabel("Message", 15, 161)
Global $msgCheckBox = GUICtrlCreateCheckbox("PreMsg", 70, 158)
$emailMsg = GUICtrlCreateInput("", 10, 180, 150, 20)

Global $autoSendCheckBox = GUICtrlCreateCheckbox("AutoSend", 10, 205)

GUISwitch($hMainGUI)
GUISetState(@SW_SHOW)

Local $aMsg = 0

While 1

    $aMsg = GUIGetMsg(1)

    Select

	  Case $aMsg[0] = $buildRateConButton
            Call("BuildRateCon")

	  Case $aMsg[0] = $emailTemplateButton
            Call("EmailRate")

	  Case $aMsg[0] = $GUI_EVENT_CLOSE And $aMsg[1] = $hMainGUI

            ExitLoop

   EndSelect

 WEnd