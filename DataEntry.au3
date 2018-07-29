#include<ExcelMod.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>
#include <String.au3>

HotKeySet("{F2}", "Terminate")
HotKeySet("^`", "BuildLoad")
HotKeySet("^1", "CollectInfo")
HotKeySet("^2", "FillInfo")

OnAutoItExitRegister ( "CloseExcelSheets" )

Global $filePath = "C:\Example\Loads.xlsx"
Global $filePath1 = "C:\Example\Drivers.xlsx"
Global $filePath2 = "C:\Example\Trucks.xlsx"


Global $oExcel = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Error", "Error creating the Excel application object.")

Global $oWorkBook = _Excel_bookOpen($oExcel, $filePath)
If @error Then
    MsgBox($MB_SYSTEMMODAL, "Error", "Error opening workbook '" & @ScriptDir & "'.")
    _Excel_Close($oExcel)
    Exit
EndIf
WinSetState("Loads", "", @SW_MINIMIZE)



Global $oExcel1 = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Error", "Error creating the Excel application object.")

Global $oWorkBook1 = _Excel_bookOpen($oExcel1, $filePath1)
If @error Then
   MsgBox($MB_SYSTEMMODAL, "Error", "Error opening workbook '" & @ScriptDir & "'.")
   _Excel_Close($oExcel)
   Exit
EndIf

Global $xDriver = _Excel_RangeRead($oWorkBook1, 1, $oWorkBook1.ActiveSheet.Usedrange.Columns("D:D"))
WinSetState("Drivers", "", @SW_MINIMIZE)



Global $oExcel2 = _Excel_Open()
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Error", "Error creating the Excel application object.")

Global $oWorkBook2 = _Excel_bookOpen($oExcel2, $filePath2)
If @error Then
   MsgBox($MB_SYSTEMMODAL, "Error", "Error opening workbook '" & @ScriptDir & "'.")
   _Excel_Close($oExcel)
   Exit
EndIf
WinSetState("Trucks", "", @SW_MINIMIZE)



Global $tabDelay = 50
Global $initialDelay = 3000
Global $newActiveLoadDelay = 2100
Global $locDelay = 2000

Global $cellMoveDelay = 350

Global $dispatcher
Global $driver
Global $truckNum
Global $shipperTimeDwnCnt
Global $consigneeTimeDwnCnt

Global $woNum
Global $shipDate
Global $delDate
Global $shipper
Global $shipperCity
Global $shipperSt
Global $PONum
Global $consignee
Global $consigneeCity
Global $consigneeSt
Global $carrierFee

Global $xLoadArray
Global $xDriverArray
Global $xTruckArray

Global $aDriverPos
Global $aTruckNumPos

Global $foundDriver
Global $foundTruck

;Func InitVars() ;Initializes all variable to default values, to reset for each load.

;Select initial Lean ID Cell and start func
Func CollectInfo()

   ;Opens and checks for CloseLoads Window
   If WinActivate("LOADBOARD") = 0 Then
	 MsgBox(0, "Window Error", "Unable to access ITS DISPATCH")
	 Return

   EndIf

   ;-----------------Work Order Number-----------------
   Send("^c")
   Sleep($cellMoveDelay)
   $woNum = ClipGet()
   Sleep($cellMoveDelay)

   $xLoadArray = _Excel_RangeFind($oWorkBook, StringStripWS($woNum, 8), Default, Default, Default, False)

   If _ArrayToString($xLoadArray) == "" Then			;Could not find load error
	  MsgBox(0, "Info", "Could not find load: " & $woNum)
	  Exit
   EndIf

   ;-----------------Dispatcher-----------------
   Send("{LEFT}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)
   $dispatcher = ClipGet()
   Sleep($cellMoveDelay)


   Switch $dispatcher ;Decides which string to use for dispatcher

	  Case "MM"
		 $dispatcher = "MOE"

	  Case "ML"
		 $dispatcher = "MICHAEL"

	  Case "MAX"
		 $dispatcher = "MAX"

	  Case "JAY"
		 $dispatcher = "JOHN"

	  Case "KA"
		 $dispatcher = "KYLE"

	  Case "SF"
		 $dispatcher = "SEAN"

	  Case "BT"
		 $dispatcher = "BILL"

	  Case "SJ"
		 $dispatcher = "JASON"

	  Case "ZZ"
		 $dispatcher = "ZEE"

	  Case "JA"
		 $dispatcher = "JAMES"

	  Case "MR"
		 $dispatcher = "MURDOCK"

	  Case "MJ"
		 $dispatcher = "MITCHELL"

	  Case "NR"
		 $dispatcher = "NICK"

	  Case "SAM"
		 $dispatcher = "SAM"

	  Case "DK"
		 $dispatcher = "DARELL"

	  Case "ISK"
		 $dispatcher = "ISMAEL"

	  Case "IS"
		 $dispatcher = "IZZY"

	  Case "DMX"
		 $dispatcher = "MAXIM"

	  Case "MxD"
		 $dispatcher = "MAXIM"

	  Case "FR"
		 $dispatcher = "FRANK"

	  Case Else
		 $dispatcher = ""

   EndSwitch

   ;-----------------Driver-----------------
   Send("{TAB 2}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)
   $driver = ClipGet()

   ;Finds white space in name and takes everything to the left including last initial.
   $driver = StringLeft($driver, (StringInStr($driver, " ") + 1))

   $xDriverArray = _Excel_RangeFind($oWorkBook1, $driver, Default, Default, Default, False)

   If _ArrayToString($xDriverArray) == "" Then			;Could not find driver error
	  MsgBox(0, "Info", "Could not find driver: " & $driver, .7)
	  $foundDriver = False
   Else
	  $aDriverPos = _StringExplode($xDriverArray[0][2], "$")
	  $foundDriver = True
   EndIf

   ;-----------------Truck Number-----------------
   Send("{TAB}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)
   $truckNum = ClipGet()

   $truckNum = StringStripWS($truckNum, 8)

   ;Removes "pup" word from truck number
   If StringInStr($truckNum, "-PUP") Then

	  $truckNum = StringLeft($truckNum, (StringInStr($truckNum, "-PUP") - 1))

   ElseIf StringInStr($truckNum, "PUP") Then

	  $truckNum = StringLeft($truckNum, (StringInStr($truckNum, "PUP") - 1))

   EndIf


   $xTruckArray = _Excel_RangeFind($oWorkBook2, $truckNum, Default, Default, Default, False)

   If _ArrayToString($xTruckArray) == "" Then			;Could not find driver error
	  MsgBox(0, "Info", "Could not find truck: " & $truckNum, .7)
	  $foundTruck = False
   Else
	  $aTruckNumPos = _StringExplode($xTruckArray[0][2], "$")
	  $foundTruck = True
   EndIf

   ;-----------------Shipper Time-----------------
   Send("{TAB 11}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)

   If ClipGet() = @CRLF Or StringInStr(ClipGet(), "ETA") > 0 Or StringIsAlpha(ClipGet()) Or StringInStr(ClipGet(), ">") Or StringInStr(ClipGet(), "/") Then		;If cell is empty or invalid, sets time to 0
	  MsgBox(0, "Info", "Invalid Pickup Time", .7)
	  $shipperTimeDwnCnt = 0
   Else
	  Local $pickTimesArray = _StringExplode( StringStripWS(ClipGet(), 8) , "-", 0)
	  Local $Array1 = _StringExplode($pickTimesArray[0], ":", 0)
	  $shipperTimeDwnCnt = ($Array1[0] * 4) + ($Array1[1] / 15)
   EndIf

   ;-----------------Consignee Time-----------------
   Send("{TAB}")
   Sleep($cellMoveDelay)
   Send("^c")
   Sleep($cellMoveDelay)
   $consigneeTime = ClipGet()
   Sleep($cellMoveDelay)

   If ClipGet() = @CRLF Or StringInStr(ClipGet(), "ETA") > 0 Or StringIsAlpha(ClipGet()) Or StringInStr(ClipGet(), ">") Or StringInStr(ClipGet(), "/") Then		;If cell is empty or invalid, sets time to 0
	  MsgBox(0, "Info", "Invalid Drop Time", .7)
	  $consigneeTimeDwnCnt = 0
   Else
	  Local $dropTimesArray = _StringExplode( StringStripWS(ClipGet(), 8) , "-", 0)
	  Local $Array2 = _StringExplode($dropTimesArray[0], ":", 0)
	  $consigneeTimeDwnCnt = ($Array2[0] * 4) + ($Array2[1] / 15)
   EndIf


   Local $aLoadPos = _StringExplode($xLoadArray[0][2], "$")

   $shipDate		= _Excel_RangeRead($oWorkbook, Default, "D" & $aLoadPos[2], 3)			;<======Adjust Columns Here
   $delDate			= _Excel_RangeRead($oWorkbook, Default, "E" & $aLoadPos[2], 3)
   $shipper			= _Excel_RangeRead($oWorkbook, Default, "G" & $aLoadPos[2], 3)
   $shipperCity		= _Excel_RangeRead($oWorkbook, Default, "H" & $aLoadPos[2], 3)
   $shipperSt		= _Excel_RangeRead($oWorkbook, Default, "I" & $aLoadPos[2], 3)
   $PONum			= _Excel_RangeRead($oWorkbook, Default, "J" & $aLoadPos[2], 3)
   $consignee		= _Excel_RangeRead($oWorkbook, Default, "K" & $aLoadPos[2], 3)
   $consigneeCity	= _Excel_RangeRead($oWorkbook, Default, "L" & $aLoadPos[2], 3)
   $consigneeSt		= _Excel_RangeRead($oWorkbook, Default, "M" & $aLoadPos[2], 3)
   $carrierFee		= _Excel_RangeRead($oWorkbook, Default, "P" & $aLoadPos[2], 3)

   ;Formats PONumber to take only first number in cell
   $PONum = StringStripWS($PONum, 8)
   If StringInStr($PONum, "/") Then
	  $PONum = StringLeft($PONum, (StringInStr($PONum, "/") - 1))
   EndIf

   ConsoleWrite("WO Number: " & $woNum & " Ship Date: " & $shipDate & " Del Date: " & $delDate & " Shipper: " & $shipper & " PONum: " & $PONum & " Consignee: " & $consignee & " Carrier Fee: " & $carrierFee & " Dispatcher: " & $dispatcher)
   ConsoleWrite(" Driver: " & $driver)


EndFunc

Func FillInfo()

   WinActivate("ITS DISPATCH")

   ;-----------------New Active Load-----------------
   Send("{TAB 22}")									;<======Adjust Tabs Here

   Send("{enter}")

   Sleep($newActiveLoadDelay)

   ;----------------- Delivery Notes-----------------
   Send("+{TAB 34}")									;<======Adjust Tabs Here
   Send($consignee & " " & $consigneeCity & ", " & $consigneeSt)

   ;-----------------CONSIGNEE TIME-----------------
   Send("+{TAB 6}")

   For $i = 1 To $consigneeTimeDwnCnt Step +1

	  Send("{DOWN}")

   Next

   ;-----------------CONSIGNEE DATE-----------------
   Send("+{TAB}")
   ClipPut($delDate)
   Send("{APPSKEY}")
   Sleep(50)
   Send("{DOWN 3}")
   Send("{ENTER}")

   ;-----------------CONSIGNEE-----------------
   Send("+{TAB 3}")
   Send($consignee)
   Sleep($locDelay)

   ;-----------------SHIPPING NOTES-----------------
   Send("+{TAB 7}")
   Send($shipper & " " & $shipperCity & ", " & $shipperSt)


   ;-----------------SHIPPER TIME-----------------
   Send("+{TAB 8}")

   For $i = 1 To $shipperTimeDwnCnt Step +1

	  Send("{DOWN}")

   Next

   ;-----------------SHIP DATE-----------------
   Send("+{TAB}")
   ClipPut($shipDate)
   Send("{APPSKEY}")
   Sleep(50)
   Send("{DOWN 3}")
   Send("{ENTER}")

   ;-----------------SHIPPER-----------------
   Send("+{TAB 3}")
   Send($shipper)
   Sleep($locDelay)

   ;-----------------FLAT RATE-----------------
   Local $flatRate =  $carrierFee * .77
   Send("+{TAB 5}")
   Send($flatRate)

   ;-----------------Truck-----------------
   Send("+{TAB 2}")

   If $foundTruck = True Then
	  Send("{ENTER}")

	  For $i = 1 To ($aTruckNumPos[2] - 1) Step +1

		 Send("{DOWN}")

	  Next

	  Send("{ENTER}")
	  Sleep(100)
	  Send("{ESC}")

   EndIf

   ;-----------------EQUIPMENT TYPE-----------------
   Send("+{TAB}")

   Send("{DOWN 4}")

   ;-----------------DRIVER-----------------
   Send("+{TAB}")

   If $foundDriver = True Then
	  For $i = 1 To ($aDriverPos[2] + 3) Step +1

		 Send("{DOWN}")

	  Next

   EndIf
   ;-----------------CARRIER FEE-----------------
   Send("+{TAB 13}")
   Send($carrierFee)

   ;-----------------LEAN ID-----------------
   Send("+{TAB 4}")
   Send($woNum)

   ;-----------------BILL TO-----------------
   Send("+{TAB 4}")
   Send("dm world logistics llc")

   Sleep(1500)

   ;-----------------Numbers tab-----------------
   Send("+{TAB 2}")

   Send("{ENTER}")

   Send("{TAB 73}")
   Send($carrierFee)

   Send("{TAB 18}")
   Send($dispatcher)

   Send("{TAB 9}")

   Send("DMW")

   Send("{TAB 9}")

   Send("DNI-" & $PONum)

EndFunc


Func BuildLoad()

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

   ClipPut($woNum)

EndFunc

Func FormatStr($sStr)

   ;Takes the string before the give character.
   If StringInStr($sStr, "(") Then
	  $sStr = StringLeft($sStr, (StringInStr($sStr, "(") - 1))
   EndIf

   If StringInStr($sStr, "/") Then
	  $sStr = StringLeft($sStr, (StringInStr($sStr, "/") - 1))
   EndIf

   If StringInStr($sStr, "-") Then
	  $sStr = StringLeft($sStr, (StringInStr($sStr, "-") - 1))
   EndIf

   $sStr = StringUpper($sStr)

   $sStr = StringStripWS($sStr, 8)

   Return $sStr

EndFunc

;-----------------Terminate Function-----------------
;To terminate when esc is pressed
Func Terminate()

   MsgBox(0, "Info", "Closing...", .5)
   Exit

EndFunc

Func CloseExcelSheets()
   WinClose("Loads")
   WinClose("Trucks")
   WinClose("Drivers")
EndFunc

While 1

   Sleep(50)

WEnd