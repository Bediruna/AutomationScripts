;Bedir Aygun
;Traverses spread sheet, reads, stores, and parses data. Then inputs into another window
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <String.au3>
#include <Constants.au3>
#include <StringConstants.au3>

AutoItSetOption("MouseCoordMode", 0)

HotKeySet("{ESC}", "Terminate")
HotKeySet("^`", "UpdateLoad")
HotKeySet("^1", "OpenLoad")
HotKeySet("^2", "CollectInfo")
HotKeySet("^3", "FillInfo")

Global $OrderNum           ;1
Global $DriverName         ;2
Global $TruckNum           ;3
Global $PhoneNum           ;4
Global $PickIn             ;5
Global $PickOut            ;6
Global $DropIn             ;7
Global $DropOut            ;8

Global $mousePosX = 1800
Global $mousePosY = 800

Global $cellMoveDelay = 300
Global $loadOpenDelay = 1000

;Select initial Lean ID Cell and start func
Func CollectInfo()

   WinActivate("SpreadSheet");window name

   ;Order Number
   Send("^c")
   Sleep($cellMoveDelay)
   $OrderNum = StringStripWS(ClipGet(), 8)
   Sleep($cellMoveDelay)

   ;Driver Name
   Send("{TAB}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)
   $DriverName = ClipGet()
   ;Finds white space in name and takes everything to the left including last initial.
   $DriverName = StringLeft($DriverName, (StringInStr($DriverName, " ") + 1))

   ;Truck Number
   Send("{TAB}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)
   $TruckNum = ClipGet()


   $TruckNum = StringStripWS($TruckNum, 8)

   ;Removes "pup" word from truck number
   If StringInStr($TruckNum, "-PUP") Then

	  $TruckNum = StringLeft($TruckNum, (StringInStr($TruckNum, "-PUP") - 1))

   ElseIf StringInStr($TruckNum, "PUP") Then

	  $TruckNum = StringLeft($TruckNum, (StringInStr($TruckNum, "PUP") - 1))

   EndIf

   ;Phone Number
   Send("{TAB}")
   Sleep($cellMoveDelay)
   Send("{TAB}")
   Send("^c")
   Sleep($cellMoveDelay)
   $PhoneNum = ClipGet()
   Sleep($cellMoveDelay)


   ;Pickup Times
   Send("{TAB 9}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)

   If ClipGet() == @CRLF Or StringInStr(ClipGet(), "ETA") > 0 Or StringInStr(ClipGet(), "@") > 0 Or StringIsAlpha(ClipGet()) Then

	  ;Reset Pick times.
	  $PickIn = ""
	  $PickOut = ""

   Else

	  Local $pickTimesArray = _StringExplode(ClipGet(), "-", 0)

	  $PickIn = StringStripWS($pickTimesArray[0],8)
	  $PickOut = StringStripWS($pickTimesArray[1],8)

   EndIf



   ;Drop times
   Send("{TAB}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)

   ;If del time is empty, do nothing.
   If ClipGet() == @CRLF Or StringInStr(ClipGet(), "ETA") > 0 Or StringInStr(ClipGet(), "@") > 0 Or StringIsAlpha(ClipGet()) Then

	  ;Resetting drop times for each load.
	  $DropIn = ""
	  $DropOut = ""

   Else

	  Local $dropTimesArray = _StringExplode(ClipGet(), "-", 0)

	  $DropIn = StringStripWS($dropTimesArray[0], 8)
	  $DropOut = StringStripWS($dropTimesArray[1], 8)

   EndIf

   ;ConsoleWrite( "Order Num: " & $OrderNum & " Driver Name: " & $DriverName & " Truck Number: " & $TruckNum & " Phone Number: " & $PhoneNum & " Pick In: " & $PickIn & " Pick Out: " & $PickOut & " Drop In: " & $DropIn & " Drop Out: " & $DropOut)

EndFunc

Func FillInfo()

   WinActivate("Close Loads")

   ;Click to make sure text box is not selected
   MouseClick($MOUSE_CLICK_LEFT, $mousePosX, $mousePosY, 1, 1)

   Send("/")
   Send("TMS ID(")
   Send("{TAB}")

   Send("^a")

   Send($OrderNum)

   Send("{ENTER}")
   Sleep($loadOpenDelay)

   GetAppts()

   Send("/")
   Send("AMAZON LOGI")
   Send("{TAB 4}")

   Send($DriverName & ". " & $PhoneNum)

   Send("{TAB}")

   Send($TruckNum)

   Send("{TAB 2}")

   If $PickIn Not = "" Then Send($pickApptDate)

   Send("{TAB}")
   Send($PickIn)

   Send("{TAB}")
   If $PickOut Not = "" Then Send($pickApptDate)

   Send("{TAB}")
   Send($PickOut)

   Send("{TAB 2}")
   If $DropIn Not = "" Then Send($dropApptDate)

   Send("{TAB}")
   Send($DropIn)

   Send("{TAB}")
   If $DropOut Not = "" Then Send($dropApptDate)

   Send("{TAB}")
   Send($DropOut)


EndFunc

;Internal Function to get appt times
Func GetAppts()

   Local $pickAppt

   Send("/")
   Send("Loaded By ")
   Send("+{DOWN 3}")
   Send("^c")

   ;Click to make sure text box is not selected
   MouseClick($MOUSE_CLICK_LEFT, $mousePosX, $mousePosY, 1, 1)

   $pickAppt = StringStripWS(StringTrimLeft(ClipGet(), 14), 2)

   $pickAppt = _StringExplode($pickAppt, " ", 0)

   Global $pickApptDate = $pickAppt[0]
   Global $pickApptTime = $pickAppt[1]


   Local $aPickApptDate = _StringExplode($pickApptDate, "/", 0)

   Local $pickApptHour = _StringExplode($pickApptTime, ":", 0)
   Local $pickHour = _StringExplode($PickIn, ":", 0)
   If $pickHour[0] > ($pickApptHour[0] + 1) then
	  $pickApptDate = $aPickApptDate[0] & "/" & $aPickApptDate[1] - 1 & "/" & $aPickApptDate[2]
   EndIf


  ;------------DROP------------
   Local $dropAppt

   Send("/")
   Send("Received By ")
   Send("+{DOWN 3}")
   Send("^c")

   ;Click to make sure text box is not selected
   MouseClick($MOUSE_CLICK_LEFT, $mousePosX, $mousePosY, 1, 1)

   $dropAppt = StringStripWS(StringTrimLeft(ClipGet(), 15), 2) ; Sposed to be 16 instead of 15

   $dropAppt = _StringExplode($dropAppt, " ", 0)

   Global $dropApptDate = $dropAppt[0]
   Global $dropApptTime = $dropAppt[1]


   Local $aDropApptDate = _StringExplode($dropApptDate, "/", 0)

   Local $dropApptHour = _StringExplode($dropApptTime, ":", 0)
   Local $dropHour = _StringExplode($DropIn, ":", 0)
   If $dropHour[0] > ($dropApptHour[0] + 1) then
	  $dropApptDate = $aDropApptDate[0] & "/" & $aDropApptDate[1] - 1 & "/" & $aDropApptDate[2]
   EndIf

EndFunc

Func UpdateLoad()

   CollectInfo()
   FillInfo()

EndFunc

;Copies current selection and opens load in Lean.
Func OpenLoad()

   Local $aPos = MouseGetPos()

   Send("^c")

   ;Opens and checks for CloseLoads Window
   If WinActivate("Close Loads") = 0 Then
	  MsgBox(0, "Window Error", "Unable to access Close Loads page")
	  Return

   EndIf

   ;Click to make sure text box is not selected
   MouseClick($MOUSE_CLICK_LEFT, $mousePosX, $mousePosY, 1, 1)

   Send("/")
   Send("TMS ID(")
   Send("{TAB}")

   Send("^a")

   Send("^v")

   Send("{ENTER}")

   WinActivate("SpreadSheet");window name

   MouseMove($aPos[0], $aPos[1], 0)

EndFunc

;-----------------Terminate Function-----------------
;To terminate when esc is pressed
Func Terminate()

   MsgBox(0, "UpdateBot", "Closing...", .7)
   Exit

EndFunc

While 1

   Sleep(50)

WEnd
