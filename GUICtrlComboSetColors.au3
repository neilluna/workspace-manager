#include <GuiComboBox.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <WinAPI.au3>
#include <Array.au3>
#include <WinAPITheme.au3>
#include <GDIPlus.au3>

#Region GUICtrlComboSetColors UDF

Global $g__aWM_CTLCOLORLISTBOX[1][16] = [[0, 0, 0]] ; init. the Global array

; #FUNCTION# ====================================================================================================================
; Name...........: GUICtrlComboSetColors
; Description ...: Change the colors and position/size of a ComboBox
; Syntax.........: GUICtrlComboSetColors ( $idCombo [, $iBgColor = Default] [, $iFgColor = Default] [, $iExtendLeft = Default] )
; Parameters ....: $idCombo     - GUICtrlCreateCombo() ControlID / [ ArrayIndex ]
;                  $iBgColor    - Background RGB color
;                               - or "-1" to use prior color declared
;                               - or to remove a control by ControlID, "-2"
;                               - or to remove a control by ArrayIndex, "-3"
;                  $iFgColor    - Foreground RGB color
;                               - or "-1" to use prior color declared
;                               - or "-2" to use sytem color and leave theme default
;                  $iExtendLeft - pixels to extend the dropdown list
;                               - or "-1" to use prior width declared
;                               - or "1" auto size, extending left  ( see Remarks/AutoSize )
;                               - or "2" auto size, extending right ( see Remarks/AutoSize )
; Return values .: Success      - index position in the array
;                  Failure      - 0
;                  @error       - 1 : Control handle = 0
;                               - 2 : GetComboBoxInfo failed
;                               - 3 : Control for removal not found
;                  @extended    - 2 : Success on Control removal
; Author ........: argumentum
; Modified.......: v0.0.0.5
; Remarks .......: this UDF is in its a work in progress, will expand if needed.
;     AutoSize...: use the pertinent parameters from GUICtrlComboSetColors_SetAutoSize()
;                  minus the CtrlID as semicolon separated to initialize. Ex: "2;Arial;8.5;0"
; Related .......:
; Link ..........: https://www.autoitscript.com/forum/topic/191035-combobox-set-dropdownlist-bgcolor/
; Example .......: Yes, at the end of the file
; ===============================================================================================================================
Func GUICtrlComboSetColors($idCombo = 0, $iBgColor = Default, $iFgColor = Default, $iExtendLeft = Default)
    If Not $idCombo Then Return SetError(1, 0, 0)
    Local $n, $tInfo, $i = 0
    If $iBgColor = -2 Or $iBgColor = -3 Then
        Local $m
        For $n = 1 To $g__aWM_CTLCOLORLISTBOX[0][0]
            If ($g__aWM_CTLCOLORLISTBOX[$n][0] = $idCombo And $iBgColor = -2) Or ($g__aWM_CTLCOLORLISTBOX[$n][9] = $idCombo And $iBgColor = -3) Then
                _ArrayDelete($g__aWM_CTLCOLORLISTBOX, $n)
                $g__aWM_CTLCOLORLISTBOX[0][0] -= 1
                Return SetError(0, 2, $n)
            EndIf
        Next
        Return SetError(3, 0, 0)
    EndIf
    For $n = 1 To $g__aWM_CTLCOLORLISTBOX[0][0]
        If $g__aWM_CTLCOLORLISTBOX[$n][0] = $idCombo Then
            $i = $n
            ExitLoop
        EndIf
    Next
    If Not $i Then
        $g__aWM_CTLCOLORLISTBOX[0][0] += 1
        $i = $g__aWM_CTLCOLORLISTBOX[0][0] ;
        If $i >= UBound($g__aWM_CTLCOLORLISTBOX) Then
            ReDim $g__aWM_CTLCOLORLISTBOX[$i + 1][16] ; add extra "slots"
        EndIf
    EndIf
    Local $sStr = GUICtrlRead($idCombo)
    Local $iSetWindowTheme = 1
    If $iBgColor = Default And $iFgColor = Default Then $iSetWindowTheme = 0
    If $iBgColor = Default Then $iBgColor = _WinAPI_GetSysColor($COLOR_WINDOW)
    If $iFgColor = Default Then $iFgColor = _WinAPI_GetSysColor($COLOR_WINDOWTEXT)
    If $iBgColor = -1 Then $iBgColor = $g__aWM_CTLCOLORLISTBOX[$i][10]
    If $iFgColor = -1 Then $iFgColor = $g__aWM_CTLCOLORLISTBOX[$i][11]
    $g__aWM_CTLCOLORLISTBOX[$i][11] = $iFgColor
    $g__aWM_CTLCOLORLISTBOX[$i][10] = $iBgColor

    If $iExtendLeft = Default Then
        $iExtendLeft = 0
        $g__aWM_CTLCOLORLISTBOX[$i][12] = 0
    EndIf
    If $iExtendLeft = -1 Then
        $iExtendLeft = $g__aWM_CTLCOLORLISTBOX[$i][8]
    ElseIf Int($iExtendLeft) = 1 Then
        $g__aWM_CTLCOLORLISTBOX[$i][12] = 1
    ElseIf Int($iExtendLeft) = 2 Then
        $g__aWM_CTLCOLORLISTBOX[$i][12] = 2
    Else
        $g__aWM_CTLCOLORLISTBOX[$i][12] = 0
    EndIf

    $g__aWM_CTLCOLORLISTBOX[$i][8] = Int($iExtendLeft)
    $g__aWM_CTLCOLORLISTBOX[$i][0] = $idCombo
    $g__aWM_CTLCOLORLISTBOX[$i][1] = GUICtrlGetHandle($idCombo)

    $g__aWM_CTLCOLORLISTBOX[$i][13] = "Arial" ; default $sFont
    $g__aWM_CTLCOLORLISTBOX[$i][14] = 8.5 ; default $fSize
    $g__aWM_CTLCOLORLISTBOX[$i][15] = 0 ; default $iStyle
    If $g__aWM_CTLCOLORLISTBOX[$i][12] Then
        $f = StringSplit($iExtendLeft, ";")
        If UBound($f) > 1 Then $g__aWM_CTLCOLORLISTBOX[$i][8] = Int($f[1])
        If UBound($f) > 2 Then $g__aWM_CTLCOLORLISTBOX[$i][13] = $f[2]
        If UBound($f) > 3 Then $g__aWM_CTLCOLORLISTBOX[$i][14] = Int($f[3])
        If UBound($f) > 4 Then $g__aWM_CTLCOLORLISTBOX[$i][15] = Int($f[4])
        $t = TimerInit()
        GUICtrlComboSetColors_SetAutoSize(Int("-" & $i), $g__aWM_CTLCOLORLISTBOX[$i][12], $g__aWM_CTLCOLORLISTBOX[$i][13], $g__aWM_CTLCOLORLISTBOX[$i][14], $g__aWM_CTLCOLORLISTBOX[$i][15])
        ConsoleWrite(TimerDiff($t) & @CRLF)
    EndIf

    If _GUICtrlComboBox_GetComboBoxInfo($idCombo, $tInfo) Then
        If $iSetWindowTheme Then
            If $g__aWM_CTLCOLORLISTBOX[$i][11] <> -2 Then _WinAPI_SetWindowTheme($g__aWM_CTLCOLORLISTBOX[$i][1], "", "")
            If $g__aWM_CTLCOLORLISTBOX[$i][11] <> -2 Then GUICtrlSetColor($g__aWM_CTLCOLORLISTBOX[$i][0], $iFgColor)
            GUICtrlSetBkColor($g__aWM_CTLCOLORLISTBOX[$i][0], $iBgColor)
        Else
            GUICtrlSetBkColor($g__aWM_CTLCOLORLISTBOX[$i][0], _WinAPI_GetSysColor($COLOR_HOTLIGHT))
            _WinAPI_SetWindowTheme($g__aWM_CTLCOLORLISTBOX[$i][1], 0, 0)
        EndIf
        $g__aWM_CTLCOLORLISTBOX[$i][2] = DllStructGetData($tInfo, "hCombo")
        $g__aWM_CTLCOLORLISTBOX[$i][3] = DllStructGetData($tInfo, "hEdit")
        $g__aWM_CTLCOLORLISTBOX[$i][4] = DllStructGetData($tInfo, "hList") ; this is what is colored
    Else
        $g__aWM_CTLCOLORLISTBOX[0][0] -= 1
        Return SetError(2, 0, 0)
    EndIf
    If Int($g__aWM_CTLCOLORLISTBOX[$i][5]) Then _WinAPI_DeleteObject($g__aWM_CTLCOLORLISTBOX[$i][5])
    $g__aWM_CTLCOLORLISTBOX[$i][5] = 0 ; holder for "_WinAPI_CreateSolidBrush()" return value
    $g__aWM_CTLCOLORLISTBOX[$i][6] = BitOR(BitAND($iBgColor, 0x00FF00), BitShift(BitAND($iBgColor, 0x0000FF), -16), BitShift(BitAND($iBgColor, 0xFF0000), 16))
    If $g__aWM_CTLCOLORLISTBOX[$i][11] = -2 Then $iFgColor = _WinAPI_GetSysColor($COLOR_WINDOWTEXT)
    $g__aWM_CTLCOLORLISTBOX[$i][7] = BitOR(BitAND($iFgColor, 0x00FF00), BitShift(BitAND($iFgColor, 0x0000FF), -16), BitShift(BitAND($iFgColor, 0xFF0000), 16))
    If Not $g__aWM_CTLCOLORLISTBOX[0][1] Then
        If $g__aWM_CTLCOLORLISTBOX[$i][4] Then
            $g__aWM_CTLCOLORLISTBOX[0][1] = GUIRegisterMsg($WM_CTLCOLORLISTBOX, "UDF_WM_CTLCOLORLISTBOX")
            If $g__aWM_CTLCOLORLISTBOX[0][1] Then OnAutoItExitRegister("OnAutoItExit_UDF_WM_CTLCOLORLISTBOX")
        EndIf
    EndIf
    $g__aWM_CTLCOLORLISTBOX[0][2] += 1
    $g__aWM_CTLCOLORLISTBOX[$i][9] = $g__aWM_CTLCOLORLISTBOX[0][2] ; internal ID
    $g__aWM_CTLCOLORLISTBOX[0][3] = TimerInit() ; to use in UDF_WM_CTLCOLORLISTBOX()
    $g__aWM_CTLCOLORLISTBOX[0][4] = 0 ; to use in UDF_WM_CTLCOLORLISTBOX()
    If $sStr Then GUICtrlSetData($idCombo, $sStr)
    Return SetError(0, 0, $g__aWM_CTLCOLORLISTBOX[0][2])
EndFunc   ;==>GUICtrlComboSetColors

Func UDF_WM_CTLCOLORLISTBOX($hWnd, $Msg, $wParam, $lParam)
    ConsoleWrite('+ Func UDF_WM_CTLCOLORLISTBOX(' & $hWnd & ', ' & $Msg & ', ' & $wParam & ', ' & $lParam & ')' & @CRLF)
    For $i = 1 To $g__aWM_CTLCOLORLISTBOX[0][0]
        If $g__aWM_CTLCOLORLISTBOX[$i][4] = $lParam Then
            If TimerDiff($g__aWM_CTLCOLORLISTBOX[0][3]) > 500 Or $g__aWM_CTLCOLORLISTBOX[0][4] <> $lParam Then
                If $g__aWM_CTLCOLORLISTBOX[$i][12] Then GUICtrlComboSetColors_SetAutoSize("-" & $i)
            EndIf
            $g__aWM_CTLCOLORLISTBOX[0][3] = TimerInit()
            $g__aWM_CTLCOLORLISTBOX[0][4] = $lParam
            If $g__aWM_CTLCOLORLISTBOX[$i][8] > 0 Then
                Local $aWPos = WinGetPos($g__aWM_CTLCOLORLISTBOX[$i][2])
                WinMove($lParam, "", $aWPos[0] - $g__aWM_CTLCOLORLISTBOX[$i][8], $aWPos[1] + $aWPos[3], $aWPos[2] + $g__aWM_CTLCOLORLISTBOX[$i][8])
            ElseIf $g__aWM_CTLCOLORLISTBOX[$i][8] < 0 Then
                Local $aWPos = WinGetPos($g__aWM_CTLCOLORLISTBOX[$i][2])
                WinMove($lParam, "", $aWPos[0], $aWPos[1] + $aWPos[3], $aWPos[2] - $g__aWM_CTLCOLORLISTBOX[$i][8])
            EndIf
            If $g__aWM_CTLCOLORLISTBOX[$i][7] >= 0 Then
                _WinAPI_SetTextColor($wParam, $g__aWM_CTLCOLORLISTBOX[$i][7])
            EndIf
            If $g__aWM_CTLCOLORLISTBOX[$i][6] >= 0 Then
                _WinAPI_SetBkColor($wParam, $g__aWM_CTLCOLORLISTBOX[$i][6])
                If Not $g__aWM_CTLCOLORLISTBOX[$i][5] Then $g__aWM_CTLCOLORLISTBOX[$i][5] = _WinAPI_CreateSolidBrush($g__aWM_CTLCOLORLISTBOX[$i][6])
                Return $g__aWM_CTLCOLORLISTBOX[$i][5]
            EndIf
            Return 0
        EndIf
    Next
EndFunc   ;==>UDF_WM_CTLCOLORLISTBOX

; #FUNCTION# ====================================================================================================================
; Name...........: GUICtrlComboSetColors_SetAutoSize
; Description ...: Set autosize for a ComboBox initialized in GUICtrlComboSetColors()
; Syntax.........: GUICtrlComboSetColors ( $idCombo [, $iExtendLeft = Default] [, $sFont = Default] [, $fSize = Default] [, $iStyle = Default] )
; Parameters ....: $idCombo     - GUICtrlCreateCombo() ControlID / [ ArrayIndex ]
;                  $iExtendLeft - 1 = Left, 2 = Right, 0 = disable auto-sizing
;                  $sFont       - Font name
;                  $fSize       - Font size
;                  $iStyle      - Font style
; Return values .: Success      - widthest string in pixels
;                  Failure      - -1
;                  @error       - look at the comments in the function
; Author ........: argumentum
; Modified.......: v0.0.0.5
; Remarks .......: this UDF is in its a work in progress, will expand if needed.
; Related .......: GUICtrlComboSetColors()
; Link ..........: https://www.autoitscript.com/forum/topic/191035-combobox-set-dropdownlist-bgcolor/
; Example .......: Yes, at the end of the file
; ===============================================================================================================================
Func GUICtrlComboSetColors_SetAutoSize($idCombo, $iExtendLeft = Default, $sFont = Default, $fSize = Default, $iStyle = Default)
    ConsoleWrite('+ Func GUICtrlComboSetColors_AutoSizeSet("' & $idCombo & '", "' & $iExtendLeft & '", "' & $sFont & '", "' & $fSize & '", "' & $iStyle & '")' & @CRLF)
    $idCombo = Int($idCombo) ; just in case the value is a string
    Local $n, $iArrayIndex = 0, $iCtrl = 0
    If $idCombo > 0 Then
        For $n = 1 To $g__aWM_CTLCOLORLISTBOX[0][0]
            If $g__aWM_CTLCOLORLISTBOX[$n][0] = $idCombo Then ; the expected value, is the ControlID
                $iArrayIndex = $n
                ExitLoop
            EndIf
        Next
        Return SetError(4, 0, -1) ; $iArrayIndex not found
    ElseIf $idCombo < 0 Then ; the expected value, is a negative of array's index ..
        $iArrayIndex = Int(StringTrimLeft(StringStripWS($idCombo, 8), 1)) ; .. so now is a positive value ..
        If $iArrayIndex < 1 Then Return SetError(3, 0, -1) ; .. else, error ..
        If $iArrayIndex > $g__aWM_CTLCOLORLISTBOX[0][0] Then Return SetError(2, 0, -1) ; .. as long as is not greater than expected
    Else
        Return SetError(1, 0, -1) ; could not find a usable value
    EndIf
    Switch $iExtendLeft
        Case 0, 1, 2
            $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][12] = $iExtendLeft
    EndSwitch
    Local $aCtrlPos = WinGetPos($g__aWM_CTLCOLORLISTBOX[$iArrayIndex][1])
    If UBound($aCtrlPos) <> 4 Then Return SetError(5, 0, -1) ; could not get a usable value
    Local $sString = StringReplace(_GUICtrlComboBox_GetList($g__aWM_CTLCOLORLISTBOX[$iArrayIndex][0]), "|", @CRLF)
    Local $aStrWidth = _GDIPlus_MeasureString($sString, $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][13], $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][14], $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][15])
    If UBound($aStrWidth) <> 2 Then Return SetError(6, 0, -1) ; could not get a usable value
    If $aStrWidth[0] < $aCtrlPos[2] Then
        $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][8] = 0
    Else
        $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][8] = $aStrWidth[0] - $aCtrlPos[2]
        If $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][12] = 2 Then $g__aWM_CTLCOLORLISTBOX[$iArrayIndex][8] = Int("-" & $aStrWidth[0] - $aCtrlPos[2])
    EndIf
    Return $aStrWidth[0]
EndFunc   ;==>GUICtrlComboSetColors_SetAutoSize


Func _GDIPlus_MeasureString($sString, $sFont = "Arial", $fSize = 12, $iStyle = 0, $bRound = True)
    ConsoleWrite('Func _GDIPlus_MeasureString("' & $sString & '", "' & $sFont & '", "' & $fSize & '", "' & $iStyle & '", "' & $bRound & '")' & @CRLF)
    ; original code @ https://www.autoitscript.com/forum/topic/150736-gdi-wrapping-text/?do=findComment&comment=1077210

    If Not $__g_iGDIPRef Then _GDIPlus_Startup() ; added by argumentum for this UDF's implementation ( AutoIt v3.3.14 ) due to the way the function is written
;~      Func _GDIPlus_Startup($sGDIPDLL = Default, $bRetDllHandle = False)
;~          $__g_iGDIPRef += 1 <-- I believe this aspect should be coded differently in "GDIPlus.au3"
;~          If $__g_iGDIPRef > 1 Then Return True


    Local $aSize[2]
    Local Const $hFamily = _GDIPlus_FontFamilyCreate($sFont)
    If Not $hFamily Then Return SetError(1, 0, $aSize)
    Local Const $hFormat = _GDIPlus_StringFormatCreate()
    Local Const $hFont = _GDIPlus_FontCreate($hFamily, $fSize, $iStyle)
    Local Const $tLayout = _GDIPlus_RectFCreate(0, 0, 0, 0)
    Local Const $hGraphic = _GDIPlus_GraphicsCreateFromHWND(0)
    Local $aInfo = _GDIPlus_GraphicsMeasureString($hGraphic, $sString, $hFont, $tLayout, $hFormat)
    $aSize[0] = $bRound ? Round($aInfo[0].Width, 0) : $aInfo[0].Width
    $aSize[1] = $bRound ? Round($aInfo[0].Height, 0) : $aInfo[0].Height
    _GDIPlus_FontDispose($hFont)
    _GDIPlus_FontFamilyDispose($hFamily)
    _GDIPlus_StringFormatDispose($hFormat)
    _GDIPlus_GraphicsDispose($hGraphic)
    Return $aSize
EndFunc   ;==>_GDIPlus_MeasureString

Func OnAutoItExit_UDF_WM_CTLCOLORLISTBOX()
    For $i = 1 To $g__aWM_CTLCOLORLISTBOX[0][0]
        If Int($g__aWM_CTLCOLORLISTBOX[$i][5]) Then _WinAPI_DeleteObject($g__aWM_CTLCOLORLISTBOX[$i][5])
    Next
    If $__g_iGDIPRef Then _GDIPlus_Shutdown()
EndFunc   ;==>OnAutoItExit_UDF_WM_CTLCOLORLISTBOX

#EndRegion GUICtrlComboSetColors UDF


; Example()

Func Example()

    ; Create GUI
    GUICreate("ComboBox Set DROPDOWNLIST BgColor", 640, 300)

    Local $a_idCombo[7] = [6]

    $a_idCombo[1] = GUICtrlCreateCombo("", 2, 2, 390, 296, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
    GUICtrlComboSetColors($a_idCombo[1], 0xEEEEEE, -2, Default)
    Example_FillTheCombo($a_idCombo[1])
    GUICtrlCreateLabel("<<< change BG color, default theme && size ", 400, 4, 396, 296)

    $a_idCombo[2] = GUICtrlCreateCombo("", 2, 32, 390, 296, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
    GUICtrlComboSetColors($a_idCombo[2], 0x0000FF, 0xFFFF00, 0)
    Example_FillTheCombo($a_idCombo[2])
    GUICtrlCreateLabel("<<< change colors", 400, 34, 396, 296)

    $a_idCombo[3] = GUICtrlCreateCombo("", 2, 62, 390, 296, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
    GUICtrlComboSetColors($a_idCombo[3], 0xdddddd, Default, 100)
    Example_FillTheCombo($a_idCombo[3])
    GUICtrlCreateLabel("<<< change BG color, resize 100px. left", 400, 64, 396, 296)

    $a_idCombo[4] = GUICtrlCreateCombo("", 2, 92, 390, 296, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
    GUICtrlComboSetColors($a_idCombo[4], Default, 0x0000FF, -100)
    Example_FillTheCombo($a_idCombo[4])
    GUICtrlCreateLabel("<<< change FG color, resize 100px. right", 400, 94, 396, 296)

    $a_idCombo[5] = GUICtrlCreateCombo("", 2, 122, 390, 296, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
    GUICtrlComboSetColors($a_idCombo[5], 0x00FFFF, 0x0000FF, 1)
    Example_FillTheCombo($a_idCombo[5])
    GUICtrlCreateLabel("<<< change colors, resize auto left", 400, 124, 396, 296)

    $a_idCombo[6] = GUICtrlCreateCombo("", 2, 152, 390, 296, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
    GUICtrlSetFont($a_idCombo[6], 10, 400, 0, "Courier New")
    GUICtrlComboSetColors($a_idCombo[6], Default, Default, "2;Courier New;10")
    Example_FillTheCombo($a_idCombo[6])
    GUICtrlCreateLabel("<<< default colors, resize auto right", 400, 154, 396, 296)

    Local $bttnArrayShow = GUICtrlCreateButton("Show array", 2, 296 - 27, 75, 25)
    Local $bttnStrMore = GUICtrlCreateButton("Longer str.", 102, 296 - 27, 75, 25)
    Local $idLorem = GUICtrlCreateLabel("", 195, 296 - 27, 50, 25)
    Local $bttnStrLess = GUICtrlCreateButton("Shorter str.", 252, 296 - 27, 75, 25)


    GUISetState(@SW_SHOW)
    WinActivate("ComboBox Set DROPDOWNLIST BgColor")

;~  Sleep(3500) ; you can reassign colors, size, or restore default
;~  GUICtrlComboSetColors($idCombo5, Default, Default, 300) ; this resets the Control back to default and changes $iExtendLeft
;~  GUICtrlComboSetColors($idCombo5, 0x0000FF, 0x00FFFF, -1) ; this changes the colors and keeps $iExtendLeft as it was
;~  GUICtrlComboSetColors($idCombo5, -1, -1, 300) ; using "-1" will keep the existing colors
;~                                              ; so in this case, only the $iExtendLeft is declared
;~  Example_FillTheCombo($idCombo5)


;~  Sleep(500) ; after removal, it will not repaint "hList", but then again, you're deleteing the control
;~  GUICtrlComboSetColors($idColors, -3)
;~  GUICtrlDelete($idCombo2)

    Local $iLorem = 5, $sLorem = ""
    Example_LoremStr($iLorem, $sLorem, $a_idCombo, $idLorem)
    ; Loop until the user exits.
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                GUIDelete()
                Return
            Case $bttnArrayShow
                _ArrayDisplay($g__aWM_CTLCOLORLISTBOX, "$g__aWM_CTLCOLORLISTBOX")
            Case $bttnStrMore
                $iLorem += 5
                Example_LoremStr($iLorem, $sLorem, $a_idCombo, $idLorem)
            Case $bttnStrLess
                $iLorem -= 5
                Example_LoremStr($iLorem, $sLorem, $a_idCombo, $idLorem)

        EndSwitch
    WEnd

EndFunc   ;==>Example

Func Example_FillTheCombo(ByRef $idComboCtrl)
    GUICtrlSetData($idComboCtrl, "")
    _GUICtrlComboBox_AddString($idComboCtrl, "something")
    _GUICtrlComboBox_AddString($idComboCtrl, "something else")
    _GUICtrlComboBox_AddString($idComboCtrl, "blah, blah, blah, blah")
    _GUICtrlComboBox_AddString($idComboCtrl, "Lorem will change")
    Local $a = _GUICtrlComboBox_GetListArray($idComboCtrl)
    GUICtrlSetData($idComboCtrl, $a[1])
EndFunc   ;==>Example_FillTheCombo

Func Example_LoremStr(ByRef $iLorem, ByRef $sLorem, ByRef $a_idCombo, ByRef $idLorem)
    Local Static $s = "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam erat volutpat."
    $s &= " Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex ea commodo consequat."
    Local Static $i = 5, $a = StringSplit($s, " ")
    If $iLorem < 1 Then $iLorem = 1
    If $iLorem > $a[0] Then $iLorem = $a[0]
    Local $x, $iLastEntry
    $sLorem = ""
    GUICtrlSetData($idLorem, $iLorem & ' words')
    For $x = 1 To $iLorem
        $sLorem &= $a[$x] & " "
    Next
    For $x = 1 To $a_idCombo[0]
        $iLastEntry = _GUICtrlComboBox_GetCount($a_idCombo[$x]) - 1
        _GUICtrlComboBox_DeleteString($a_idCombo[$x], $iLastEntry)
        _GUICtrlComboBox_AddString($a_idCombo[$x], $sLorem)
    Next
EndFunc   ;==>Example_LoremStr