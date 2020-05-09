#include <GUIConstantsEx.au3>
#include <ComboConstants.au3>
$PrnTXT = @AppDataDir & '\DefPrinter.txt'

if $CmdLine[0]>=1 Then
   if StringLower($CmdLine[1])='/config' Then
	  $wbemFlagReturnImmediately = 0x10
	  $wbemFlagForwardOnly = 0x20
	  $PrnList=""
	  $strDefaultPrinter = ""
	  $objWMIService = ObjGet("winmgmts:\\.\root\CIMV2")
	  $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_Printer", "WQL", $wbemFlagReturnImmediately + $wbemFlagForwardOnly)
	  If IsObj($colItems) then
		 For $objItem In $colItems
			$PrnList &= '|' & $objItem.DeviceID
			if $objItem.Default Then $strDefaultPrinter = $objItem.DeviceID
		 Next
	  Endif
	  ShowWin()
   EndIf
Else
   If FileExists($PrnTXT) Then
	  $Fl = FileOpen($PrnTXT,0)
	  SetDefault(FileReadLine($Fl,1))
   EndIf
EndIf

Func ShowWin()
   Local $msg
   GUICreate("Вибір принтера",370,80)
   GUICtrlCreateLabel("Виберіть, будь-ласка, принтер за замовчуванням:",10,5)
   $lst = GUICtrlCreateCombo("", 10, 22,350,15,$CBS_DROPDOWNLIST)
   GUICtrlSetData($lst,$PrnList,$strDefaultPrinter)
   $btnSave = GUICtrlCreateButton("Зберегти",135,50,100,25)
   GUISetState()
   While 1
	  $msg = GUIGetMsg()
	  If $msg = $btnSave Then
		 $Fl = FileOpen($PrnTXT,2)
		 FileWriteLine($Fl,GUICtrlRead($lst))
		 ;MsgBox(0,"Msg",GUICtrlRead($lst))
		 SetDefault(GUICtrlRead($lst))
		 ExitLoop
	  EndIf
	  If $msg = $GUI_EVENT_CLOSE Then ExitLoop
   WEnd
EndFunc

Func SetDefault($Printer)
   RunWait(@ComSpec & " /c RUNDLL32 PRINTUI.DLL,PrintUIEntry /q /y /n " & '"' & $Printer & '"', "", @SW_HIDE)
EndFunc
