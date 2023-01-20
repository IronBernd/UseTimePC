#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Comment=PC - Uptime 2 Excel
#AutoIt3Wrapper_Res_Fileversion=0.3.0.11
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Field=V0.1|Excel sheet background  must be changed
#AutoIt3Wrapper_Res_Field=V0.2|Add StartDate
#AutoIt3Wrapper_Res_Field=V0.3|ExcelSheet Improvement
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

AutoItSetOption ( "MustDeclareVars" , 1  )

#include <EventLog.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>
#include <Date.au3>
#include <Excel.au3>
#include <GuiDateTimePicker.au3>

Local $bFileInstall = FileExists(@ScriptDir & "\Monate2.xlx")
Global $g_idMemo
Dim $Day1 = @YEAR & "/01/01"
Global $useTime = _DateDiff("D", $Day1, _NowCalcDate() )
if $useTime > 45 Then
	$useTime = 45
EndIf

If $bFileInstall = 0  Then
	FileInstall( ".\_Monate.xlsx", @ScriptDir & "\Monate2.xlsx")
	$useTime = _DateDiff("D", $Day1, _NowCalcDate() )
EndIf
; Create GUI
Local $hGUI = GUICreate("UseTimePC: Zeiterfassung", 400, 300)
Local $hLabel = GUICtrlCreateLabel("Bitte Start Datum Eingeben",5,10);
Local $hDTP = _GUICtrlDTP_Create($hGUI, 5, 30, 190)
Local $hChk = GUICtrlCreateCheckbox( "Use Event:Application", 5, 55,150,15)
Local $hButton = GUICtrlCreateButton("Start Import",5,75)
Local $hDOINGS = GUICtrlCreateLabel(".",5,100,150,30);
GUISetState(@SW_SHOW)

_GUICtrlDTP_SetFormat($hDTP, "dd.MM.yyyy")
Dim $start_date;
; Loop until user exits
While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                Exit 2
			Case $hButton
				 Local $a_Date = _GUICtrlDTP_GetSystemTime($hDTP);
				 $start_date = $a_Date[2] & "." & $a_Date[1] & "." & $a_Date[0] ;
				ExitLoop
        EndSwitch
    WEnd

GUICtrlSetState($hButton,$GUI_DISABLE )
GUICtrlSetData($hDOINGS,"Read events .. This tke a while");
If _IsChecked($hChk) Then
	local $dErr = RunWait("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe  ""$v2=Get-Date '" & $start_date & "'; Get-WinEvent -Oldest -FilterHashtable @{logname='Application';id=100,101}  | Select-Object ID,TimeCreated |  Where-Object { $_.TimeCreated -ge $v2 } | ft -HideTableHeaders >" & @ScriptDir & "\b.txx" & """ ",@ScriptDir,@SW_SHOW);
;	MsgBox(0,$dErr,@error & " " & $start_date)
	ReadAppEvents()
Else
	local $dErr = RunWait("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe  ""$v2=Get-Date '" & $start_date & "'; Get-WinEvent -Oldest -FilterHashtable @{logname='system';id=50037,153}  | Select-Object ID,TimeCreated |  Where-Object { $_.TimeCreated -ge $v2 } | ft -HideTableHeaders >" & @ScriptDir & "\a.txx" & """ ",@ScriptDir,@SW_SHOW);
;	MsgBox(0,$dErr,@error & " " & $start_date)
	ReadSysEvents()
EndIf

GUIDelete($hGUI)

Func ReadAppEvents()
	DIM $aSta[3]=["H","J","L"]
	DIM $aEnd[3]  =["I","K","M"]
	GUICtrlSetData($hDOINGS,"Open Excel sheet.");
	Local $oExcel = _Excel_Open()
	Local $sWorkbook = @ScriptDir & "\Monate2.xlsx"
	Local $oWorkbook = _Excel_BookOpen($oExcel,$sWorkbook);
    Local $idx1 =0
	Local $eFile = FileOpen(@ScriptDir & "\b.txx")
	Local $sLine = FileReadLine($eFile); <empty>
	DIM $ToDay = _NowDate()
;	MsgBox(0,"Heute",$ToDay)

	Local $sLine1 = ""
	Local $sLine2 = ""
	Local $aStop
	Local $aStart
	Local $oDate =""
	while $eFile
		do
			if $sLine1 = "" Then
				$sLine  = FileReadLine($eFile) ; TimeGenerated -------------  UP
				$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
			EndIf
			$aStart = StringSplit($sLine1," ") ; ID; Datum; Stunde
			if  $aStart[1] <> 100 Then
				$sLine  = FileReadLine($eFile) ; TimeGenerated -------------  UP
				$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
			Endif
		until $aStart[1] = 100
;		MsgBox(0,$aStart[2],"Start " & $aStart[3])
		if $aStart[2] = $ToDay Then ; Log File Error
			_Excel_BookSave($oWorkbook)
			_Excel_Close($oExcel)
			Return
		EndIf
;		$aStart = StringSplit($sLine1," ") ; ID; Datum; Stunde
		Do
			$sLine  = FileReadLine($eFile) ; TimeGenerated ------------- DOWN
			$sLine2 = StringStripWS($sLine, $STR_STRIPLEADING )
			$aStop  = StringSplit($sLine2," ") ; ID; Datum; Stunde
		Until $aStop[1] = 101
;		MsgBox(0,$aStop[2],"Stop" & $aStop[3])

		$sLine = FileReadLine($eFile) ; TimeGenerated -------------  UP
		$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
		Local $aBreak = StringSplit($sLine1," ")
		Local $d=0
		while  ($aBreak[2] = $aStop[2]) and ($d <=30)
;			Msgbox(0,$aBreak[2],$aStop[2])
			Local $aZStop = StringSplit($aStop[3],":")
			Local $aZBreak = StringSplit($aBreak[3],":")
			Local $d = 60*($aZBreak[1]-$aZStop[1]) + ($aZBreak[2]-$aZStop[2])
;			MsgBox(0,$sLine2,$sLine1)
			if  $d <= 30 Then
				do
					$sLine= FileReadLine($eFile) ; TimeGenerated ------------- DOWN
					$sLine2 = StringStripWS($sLine, $STR_STRIPLEADING )
					$aStop=StringSplit($sLine2," ") ; ID; Datum; Stunde
				until $aStop[1] = 101

				$sLine = FileReadLine($eFile) ; TimeGenerated ------------- UP
				$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
			EndIf
			$aBreak = StringSplit($sLine1," ")
		WEnd
;		MsgBox(0,"T: " & $aStart[2] , $aStop[2]);

		if $aStart[2] = $aStop[2] Then ; Gleiches Datum
			if $oDate <> $aStop[2] Then
				$idx1 = 0
			Endif
			Local $aDate=StringSplit($aStart[2],".")
			Local $t = StringFormat("%s: %s ... %s",$aDate[2],$aDate[1],$aStart[2])
;			MsgBox(0,$t,$sLine1)
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[2] ,"B" & $aDate[1] + 1); Datum
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[3] ,$aSta[$idx1] & $aDate[1] + 1); Start Zeit
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStop[3]  ,$aEnd[$idx1] & $aDate[1] + 1); End Zeit
			$idx1+=1
			if $idx1 >2 Then
				$idx1 = 2
			EndIf
			$oDate = $aStop[2]
		Else ; unterschiedliches Datum

			if $oDate <> $aStart[2] Then
				$idx1 = 0
			Endif
			Local $aDate=StringSplit($aStart[2],".")
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[2] ,"B" & $aDate[1] + 1); Datum
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[3] ,$aSta[$idx1] & $aDate[1] + 1); Start Zeit
			_Excel_RangeWrite($oWorkbook,$aDate[2], "23:59:59" ,$aEnd[$idx1] & $aDate[1] + 1); End Zeit

			$idx1=0;
			$aDate=StringSplit($aStop[2],".")
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStop[2] ,"B" & $aDate[1] + 1); Datum
			_Excel_RangeWrite($oWorkbook,$aDate[2], "00:00:02" ,$aSta[$idx1] & $aDate[1] + 1); Start Zeit
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStop[3]  ,$aEnd[$idx1] & $aDate[1] + 1); End Zeit
			$idx1+=1
			if $idx1 >2 Then
				$idx1 = 2
			EndIf
			$oDate = $aStop[2]
		EndIf
	Wend
	_Excel_BookSave($oWorkbook)
	_Excel_Close($oExcel)
EndFunc   ;==>Example


Func ReadSysEvents()
	DIM $aSta[3]=["H","J","L"]
	DIM $aEnd[3]  =["I","K","M"]
	GUICtrlSetData($hDOINGS,"Open Excel sheet.");

	Local $oExcel = _Excel_Open()
	Local $sWorkbook = @ScriptDir & "\Monate2.xlsx"
	Local $oWorkbook = _Excel_BookOpen($oExcel,$sWorkbook);
    Local $idx1 =0
	Local $eFile = FileOpen(@ScriptDir & "\a.txx")
	Local $sLine = FileReadLine($eFile); <empty>
	DIM $ToDay = _NowDate()
;	MsgBox(0,"Heute",$ToDay)

	Local $sLine1 = ""
	Local $sLine2 = ""
	Local $aStop
	Local $aStart
	Local $oDate =""
	while $eFile
		do
			if $sLine1 = "" Then
				$sLine  = FileReadLine($eFile) ; TimeGenerated -------------  UP
				$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
			EndIf
			$aStart = StringSplit($sLine1," ") ; ID; Datum; Stunde
;			MsgBox(0,"find start",$aStart[1]);
			if  $aStart[1] <> 153 Then
				$sLine  = FileReadLine($eFile) ; TimeGenerated -------------  UP
				$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
			Endif
		until $aStart[1] = 153
;		MsgBox(0,$aStart[2],$ToDay)
		if $aStart[2] = $ToDay Then ; Log File Error
			_Excel_BookSave($oWorkbook)
			_Excel_Close($oExcel)
			Return
		EndIf
;		$aStart = StringSplit($sLine1," ") ; ID; Datum; Stunde
		Do
			$sLine  = FileReadLine($eFile) ; TimeGenerated ------------- DOWN
			$sLine2 = StringStripWS($sLine, $STR_STRIPLEADING )
			$aStop  = StringSplit($sLine2," ") ; ID; Datum; Stunde
		Until $aStop[1] = 50037
;		MsgBox(0,$aStop[2],"Stop")

		$sLine = FileReadLine($eFile) ; TimeGenerated -------------  UP
		$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
		Local $aBreak = StringSplit($sLine1," ")
		Local $d=0
		while  ($aBreak[2] = $aStop[2]) and ($d <=30)
;			Msgbox(0,$aBreak[2],$aStop[2])
			Local $aZStop = StringSplit($aStop[3],":")
			Local $aZBreak = StringSplit($aBreak[3],":")
			Local $d = 60*($aZBreak[1]-$aZStop[1]) + ($aZBreak[2]-$aZStop[2])
;			MsgBox(0,$sLine2,$sLine1)
			if  $d <= 30 Then
				do
					$sLine= FileReadLine($eFile) ; TimeGenerated ------------- DOWN
					$sLine2 = StringStripWS($sLine, $STR_STRIPLEADING )
					$aStop=StringSplit($sLine2," ") ; ID; Datum; Stunde
				until $aStop[1] = 50037

				$sLine = FileReadLine($eFile) ; TimeGenerated ------------- UP
				$sLine1 = StringStripWS($sLine, $STR_STRIPLEADING )
			EndIf
			$aBreak = StringSplit($sLine1," ")
		WEnd
;		MsgBox(0,"T: " & $aStart[2] , $aStop[2]);

		if $aStart[2] = $aStop[2] Then ; Gleiches Datum
			if $oDate <> $aStop[2] Then
				$idx1 = 0
			Endif
			Local $aDate=StringSplit($aStart[2],".")
			Local $t = StringFormat("%s: %s ... %s",$aDate[2],$aDate[1],$aStart[2])
;			MsgBox(0,$t,$sLine1)
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[2] ,"B" & $aDate[1] + 1); Datum
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[3] ,$aSta[$idx1] & $aDate[1] + 1); Start Zeit
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStop[3]  ,$aEnd[$idx1] & $aDate[1] + 1); End Zeit
			$idx1+=1
			if $idx1 >2 Then
				$idx1 = 2
			EndIf
			$oDate = $aStop[2]
		Else ; unterschiedliches Datum

			if $oDate <> $aStart[2] Then
				$idx1 = 0
			Endif
			Local $aDate=StringSplit($aStart[2],".")
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[2] ,"B" & $aDate[1] + 1); Datum
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStart[3] ,$aSta[$idx1] & $aDate[1] + 1); Start Zeit
			_Excel_RangeWrite($oWorkbook,$aDate[2], "23:59:59" ,$aEnd[$idx1] & $aDate[1] + 1); End Zeit

			$idx1=0;
			$aDate=StringSplit($aStop[2],".")
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStop[2] ,"B" & $aDate[1] + 1); Datum
			_Excel_RangeWrite($oWorkbook,$aDate[2], "00:00:02" ,$aSta[$idx1] & $aDate[1] + 1); Start Zeit
			_Excel_RangeWrite($oWorkbook,$aDate[2], $aStop[3]  ,$aEnd[$idx1] & $aDate[1] + 1); End Zeit
			$idx1+=1
			if $idx1 >2 Then
				$idx1 = 2
			EndIf
			$oDate = $aStop[2]
		EndIf
	Wend
	_Excel_BookSave($oWorkbook)
	_Excel_Close($oExcel)
EndFunc   ;==>Example


Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked
