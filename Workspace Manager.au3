Global $title = "Workspace Manager"
Global $version = "1.4"

AutoItSetOption("MustDeclareVars", 1)

#Include <Constants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>

#include "ExtMsgBox.au3"
#include "GUICtrlComboSetColors.au3"

Opt("GUIOnEventMode", 1)
Opt("GUIEventOptions", 1)

; List of monitor coordinates and dimensions.

Global $monitor_count = 0
Global $monitor_list[1][8] ; Screen xywh and workspace xywh per monitor.

; Set the initial action variables.

Global $size_width = "No change"
Global $size_height = "No change"
Global $move_horizontal = "No change"
Global $move_vertical = "No change"

; Create the main window and all its controls.

GUICreate($title & " - " & $version, 240, 195)
GUISetBkColor($COLOR_BLACK)
GUICtrlSetDefColor($COLOR_WHITE)
GUICtrlSetDefBkColor($COLOR_BLACK)

GUICtrlCreateGroup("Move", 10, 10, 220, 70)
GUICtrlCreateLabel("Move", 20, 10)

GUICtrlCreateLabel("Horizontal", 20, 28)
Global $move_horizontal_combo = _
    GUICtrlCreateCombo("No change", 90, 25, 130, 21, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlComboSetColors($move_horizontal_combo, $COLOR_BLACK, $COLOR_WHITE)
GUICtrlSetData($move_horizontal_combo, $move_horizontal)
_GUICtrlComboBox_AddString($move_horizontal_combo, "Center")
_GUICtrlComboBox_AddString($move_horizontal_combo, "Left edge")
_GUICtrlComboBox_AddString($move_horizontal_combo, "Right edge")
_GUICtrlComboBox_AddString($move_horizontal_combo, "Left justify with ...")
_GUICtrlComboBox_AddString($move_horizontal_combo, "Right justify with ...")
_GUICtrlComboBox_AddString($move_horizontal_combo, "Stack to the left of ...")
_GUICtrlComboBox_AddString($move_horizontal_combo, "Stack to the right of ...")

GUICtrlCreateLabel("Vertical", 20, 53)
Global $move_vertical_combo = _
    GUICtrlCreateCombo("No change", 90, 50, 130, 21, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlComboSetColors($move_vertical_combo, $COLOR_BLACK, $COLOR_WHITE)
GUICtrlSetData($move_vertical_combo, $move_vertical)
_GUICtrlComboBox_AddString($move_vertical_combo, "Center")
_GUICtrlComboBox_AddString($move_vertical_combo, "Top Edge")
_GUICtrlComboBox_AddString($move_vertical_combo, "Bottom Edge")
_GUICtrlComboBox_AddString($move_vertical_combo, "Top justify with ...")
_GUICtrlComboBox_AddString($move_vertical_combo, "Bottom justify with ...")
_GUICtrlComboBox_AddString($move_vertical_combo, "Stack above ...")
_GUICtrlComboBox_AddString($move_vertical_combo, "Stack below ...")

GUICtrlCreateGroup("Size", 10, 85, 220, 70)
GUICtrlCreateLabel("Size", 20, 85)

GUICtrlCreateLabel("Width", 20, 103)
Global $size_width_combo = _
    GUICtrlCreateCombo("No change", 90, 100, 130, 21, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlComboSetColors($size_width_combo, $COLOR_BLACK, $COLOR_WHITE)
GUICtrlSetData($size_width_combo, $size_width)
_GUICtrlComboBox_AddString($size_width_combo, "480")
_GUICtrlComboBox_AddString($size_width_combo, "640")
_GUICtrlComboBox_AddString($size_width_combo, "800")
_GUICtrlComboBox_AddString($size_width_combo, "960")
_GUICtrlComboBox_AddString($size_width_combo, "1024")
_GUICtrlComboBox_AddString($size_width_combo, "1280")
_GUICtrlComboBox_AddString($size_width_combo, "1920")
_GUICtrlComboBox_AddString($size_width_combo, "2560")
_GUICtrlComboBox_AddString($size_width_combo, "Match the width of ...")
_GUICtrlComboBox_AddString($size_width_combo, "Extend to the left of ...")
_GUICtrlComboBox_AddString($size_width_combo, "Extend to the right of ...")
_GUICtrlComboBox_AddString($size_width_combo, "Extend to the right edge")

GUICtrlCreateLabel("Height", 20, 128)
Global $size_height_combo = _
    GUICtrlCreateCombo("No change", 90, 125, 130, 21, BitOR($CBS_DROPDOWNLIST, $WS_HSCROLL, $WS_VSCROLL))
GUICtrlComboSetColors($size_height_combo, $COLOR_BLACK, $COLOR_WHITE)
GUICtrlSetData($size_height_combo, $size_height)
_GUICtrlComboBox_AddString($size_height_combo, "360")
_GUICtrlComboBox_AddString($size_height_combo, "480")
_GUICtrlComboBox_AddString($size_height_combo, "600")
_GUICtrlComboBox_AddString($size_height_combo, "720")
_GUICtrlComboBox_AddString($size_height_combo, "768")
_GUICtrlComboBox_AddString($size_height_combo, "960")
_GUICtrlComboBox_AddString($size_height_combo, "1080")
_GUICtrlComboBox_AddString($size_height_combo, "1440")
_GUICtrlComboBox_AddString($size_height_combo, "Match the height of ...")
_GUICtrlComboBox_AddString($size_height_combo, "Extend to the top of ...")
_GUICtrlComboBox_AddString($size_height_combo, "Extend to the bottom of ...")
_GUICtrlComboBox_AddString($size_height_combo, "Extend to the bottom edge")

Global $apply_button = GUICtrlCreateButton("Apply", 35, 160, 75, 25)
Global $close_button = GUICtrlCreateButton("Close", 130, 160, 75, 25)

; Specify the event handlers.

GUICtrlSetOnEvent($size_width_combo, "WidthChanged")
GUICtrlSetOnEvent($size_height_combo, "HeightChanged")
GUICtrlSetOnEvent($move_horizontal_combo, "MoveHorizontalChanged")
GUICtrlSetOnEvent($move_vertical_combo, "MoveVerticalChanged")

GUICtrlSetOnEvent($apply_button, "ApplyChanges")
GUICtrlSetOnEvent($close_button, "CloseApplication")

GUISetOnEvent($GUI_EVENT_RESTORE, "RestoreApplication")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "MinimizeApplication")
GUISetOnEvent($GUI_EVENT_CLOSE, "CloseApplication")

; Show the window and wait for events.

GUISetState(@SW_SHOW)
While 1
    Sleep(1000)
WEnd

; Event handlers.

Func MoveHorizontalChanged()
    $move_horizontal = GUICtrlRead($move_horizontal_combo)
EndFunc

Func MoveVerticalChanged()
    $move_vertical = GUICtrlRead($move_vertical_combo)
EndFunc

Func WidthChanged()
    $size_width = GUICtrlRead($size_width_combo)
EndFunc

Func HeightChanged()
    $size_height = GUICtrlRead($size_height_combo)
EndFunc

Func AskForReferenceWindow($description)
    Return AskForWindow($description, "Press Ok, then select the reference window.")
EndFunc

Func AskForTargetMonitor()
    Return AskForWindow("", "Press Ok, then move your mouse to the monitor where the window will reside.")
EndFunc

Func AskForTargetWindow()
    Return AskForWindow("", "Press Ok, then select the window to size and/or move.")
EndFunc

Func AskForWindow($description, $action)
    Local $prompt = $action
    If $description <> "" Then
        $prompt = $description & @CRLF & $action
    EndIf
    _ExtMsgBoxSet(1, 0, $COLOR_BLACK, $COLOR_WHITE, -1, -1, -1, -1, "~", $COLOR_BLACK, $COLOR_WHITE)
    Local $choice = _ExtMsgBox(0, 1, $title, $prompt, 60)
    If ($choice < 0 or $choice == 2) Then
        Return 0
    EndIf
    Sleep(5000)
    Return 1
EndFunc

Func GetMonitorInfo()
    $monitor_count = 0
    Local $monitor_callback = DllCallbackRegister("MonitorEnumProc", "int", "hwnd;hwnd;ptr;lparam")
    DllCall("user32.dll", "int", "EnumDisplayMonitors", "hwnd", 0, "ptr", 0, "ptr", _
        DllCallbackGetPtr($monitor_callback), "lparam", 0)
    DllCallbackFree($monitor_callback)
EndFunc

Func MonitorEnumProc($monitor, $hdc, $rect_ptr, $lparam)
    Local $monitor_info = DllStructCreate("dword; long; long; long; long; long; long; long; long; dword")
    DllStructSetData($monitor_info, 1, DllStructGetSize($monitor_info))
    DllCall("User32.dll", "bool", "GetMonitorInfo", "hwnd", $monitor, "ptr", DllStructGetPtr($monitor_info))
    $monitor_count += 1
    ReDim $monitor_list[$monitor_count][8]
    $monitor_list[$monitor_count - 1][0] = DllStructGetData($monitor_info, 2)
    $monitor_list[$monitor_count - 1][1] = DllStructGetData($monitor_info, 3)
    $monitor_list[$monitor_count - 1][2] = DllStructGetData($monitor_info, 4) - $monitor_list[$monitor_count - 1][0]
    $monitor_list[$monitor_count - 1][3] = DllStructGetData($monitor_info, 5) - $monitor_list[$monitor_count - 1][1]
    $monitor_list[$monitor_count - 1][4] = DllStructGetData($monitor_info, 6)
    $monitor_list[$monitor_count - 1][5] = DllStructGetData($monitor_info, 7)
    $monitor_list[$monitor_count - 1][6] = DllStructGetData($monitor_info, 8) - $monitor_list[$monitor_count - 1][4]
    $monitor_list[$monitor_count - 1][7] = DllStructGetData($monitor_info, 9) - $monitor_list[$monitor_count - 1][5]
    Return 1
EndFunc

Func GetMonitorFromMouse()
    Local $mouse_pos = MouseGetPos()
    Local $monitor = 0
    For $i = 0 To $monitor_count - 1
        If ($mouse_pos[0] >= $monitor_list[$i][4]) And _
           ($mouse_pos[0] < ($monitor_list[$i][4] + $monitor_list[$i][6])) And _
           ($mouse_pos[1] >= $monitor_list[$i][5]) And _
           ($mouse_pos[1] < ($monitor_list[$i][5] + $monitor_list[$i][7])) Then
            $monitor = $i
        EndIf
    Next
    Return $monitor
EndFunc

Func ApplyChanges()
    GetMonitorInfo()

    Local $horizontal_reference_window
    If $move_horizontal = "Left justify with ..." Then
        If Not AskForReferenceWindow("Left justify with another window.") Then
            Return
        EndIf
        $horizontal_reference_window = WinGetPos("")
    ElseIf $move_horizontal = "Right justify with ..." Then
        If Not AskForReferenceWindow("Right justify with another window.") Then
            Return
        EndIf
        $horizontal_reference_window = WinGetPos("")
    ElseIf $move_horizontal = "Stack to the left of ..." Then
        If Not AskForReferenceWindow("Stack to the left of another window.") Then
            Return
        EndIf
        $horizontal_reference_window = WinGetPos("")
    ElseIf $move_horizontal = "Stack to the right of ..." Then
        If Not AskForReferenceWindow("Stack to the right of another window.") Then
            Return
        EndIf
        $horizontal_reference_window = WinGetPos("")
    EndIf

    Local $vertical_reference_window
    If $move_vertical = "Top justify with ..." Then
        If Not AskForReferenceWindow("Top justify with another window.") Then
            Return
        EndIf
        $vertical_reference_window = WinGetPos("")
    ElseIf $move_vertical = "Bottom justify with ..." Then
        If Not AskForReferenceWindow("Bottom justify with another window.") Then
            Return
        EndIf
        $vertical_reference_window = WinGetPos("")
    ElseIf $move_vertical = "Stack above ..." Then
        If Not AskForReferenceWindow("Stack above another window.") Then
            Return
        EndIf
        $vertical_reference_window = WinGetPos("")
    ElseIf $move_vertical = "Stack below ..." Then
        If Not AskForReferenceWindow("Stack below another window.") Then
            Return
        EndIf
        $vertical_reference_window = WinGetPos("")
    EndIf

    Local $width_reference_window
    If $size_width = "Match the width of ..." Then
        If Not AskForReferenceWindow("Match the width of another window.") Then
            Return
        EndIf
        $width_reference_window = WinGetPos("")
    ElseIf $size_width = "Extend to the left of ..." Then
        If Not AskForReferenceWindow("Extend the width to the left edge of another window.") Then
            Return
        EndIf
        $width_reference_window = WinGetPos("")
    ElseIf $size_width = "Extend to the right of ..." Then
        If Not AskForReferenceWindow("Extend the width to the right edge of another window.") Then
            Return
        EndIf
        $width_reference_window = WinGetPos("")
    EndIf

    Local $height_reference_window
    If $size_height = "Match the height of ..." Then
        If Not AskForReferenceWindow("Match the height of another window.") Then
            Return
        EndIf
        $height_reference_window = WinGetPos("")
    ElseIf $size_height = "Extend to the top of ..." Then
        If Not AskForReferenceWindow("Extend the height to the top edge of another window.") Then
            Return
        EndIf
        $height_reference_window = WinGetPos("")
    ElseIf $size_height = "Extend to the bottom of ..." Then
        If Not AskForReferenceWindow("Extend the height to the bottom edge of another window.") Then
            Return
        EndIf
        $height_reference_window = WinGetPos("")
    EndIf

    Local $target_monitor = 0
    If ($monitor_count > 1) And _
        (($size_width = "Extend to the right edge") Or _
         ($move_horizontal = "Center") Or _
         ($move_horizontal = "Left edge") Or _
         ($move_horizontal = "Right edge") Or _
         ($size_height = "Extend to the bottom edge") Or _
         ($move_vertical = "Center") Or _
         ($move_vertical = "Top Edge") Or _
         ($move_vertical = "Bottom Edge")) Then
        If Not AskForTargetMonitor() Then
            Return
        EndIf
        $target_monitor = GetMonitorFromMouse()
    EndIf

    If Not AskForTargetWindow() Then
        Return
    EndIf
    Local $target_window = WinGetPos("")

    Select
        Case $move_horizontal = "Center"
            $target_window[0] = $monitor_list[$target_monitor][4] + _
                ($monitor_list[$target_monitor][6] - $target_window[2]) / 2
        Case $move_horizontal = "Left edge"
            $target_window[0] = $monitor_list[$target_monitor][4]
        Case $move_horizontal = "Right edge"
            $target_window[0] = $monitor_list[$target_monitor][4] + $monitor_list[$target_monitor][6] _
            - $target_window[2]
        Case $move_horizontal = "Left justify with ..."
            $target_window[0] = $horizontal_reference_window[0]
        Case $move_horizontal = "Right justify with ..."
            $target_window[0] = $horizontal_reference_window[0] + $horizontal_reference_window[2] _
            - $target_window[2]
        Case $move_horizontal = "Stack to the left of ..."
            $target_window[0] = $horizontal_reference_window[0] - $target_window[2]
        Case $move_horizontal = "Stack to the right of ..."
            $target_window[0] = $horizontal_reference_window[0] + $horizontal_reference_window[2]
    EndSelect

    Select
        Case $move_vertical = "Center"
            $target_window[1] = $monitor_list[$target_monitor][5] + _
                ($monitor_list[$target_monitor][7] - $target_window[3]) / 2
        Case $move_vertical = "Top Edge"
            $target_window[1] = $monitor_list[$target_monitor][5]
        Case $move_vertical = "Bottom Edge"
            $target_window[1] = $monitor_list[$target_monitor][5] + $monitor_list[$target_monitor][7] _
            - $target_window[3]
        Case $move_vertical = "Top justify with ..."
            $target_window[1] = $vertical_reference_window[1]
        Case $move_vertical = "Bottom justify with ..."
            $target_window[1] = $vertical_reference_window[1] + $vertical_reference_window[3] _
            - $target_window[3]
        Case $move_vertical = "Stack above ..."
            $target_window[1] = $vertical_reference_window[1] - $target_window[3]
        Case $move_vertical = "Stack below ..."
            $target_window[1] = $vertical_reference_window[1] + $vertical_reference_window[3]
    EndSelect

    If $size_width = "Match the width of ..." Then
        $target_window[2] = $width_reference_window[2]
    ElseIf $size_width = "Extend to the left of ..." Then
        $target_window[2] = $width_reference_window[0] - $target_window[0]
    ElseIf $size_width = "Extend to the right of ..." Then
        $target_window[2] = $width_reference_window[0] + $width_reference_window[2] - $target_window[0]
    ElseIf $size_width = "Extend to the right edge" Then
        $target_window[2] = $monitor_list[$target_monitor][4] + $monitor_list[$target_monitor][6] _
        - $target_window[0]
    ElseIf StringIsInt($size_width) = 1 Then
        $target_window[2] = Int($size_width)
    EndIf

    If $size_height = "Match the height of ..." Then
        $target_window[3] = $height_reference_window[3]
    ElseIf $size_height = "Extend to the top of ..." Then
        $target_window[3] = $height_reference_window[1] - $target_window[1]
    ElseIf $size_height = "Extend to the bottom of ..." Then
        $target_window[3] = $height_reference_window[1] + $height_reference_window[3] - $target_window[1]
    ElseIf $size_height = "Extend to the bottom edge" Then
        $target_window[3] = $monitor_list[$target_monitor][5] + $monitor_list[$target_monitor][7] _
        - $target_window[1]
    ElseIf StringIsInt($size_height) = 1 Then
        $target_window[3] = Int($size_height)
    EndIf

    WinMove("", "", $target_window[0], $target_window[1], $target_window[2], $target_window[3])

    SoundPlay(@ScriptDir & "\ding.wav", 1)
EndFunc

Func RestoreApplication()
    GUISetState(@SW_RESTORE)
EndFunc

Func MinimizeApplication()
    GUISetState(@SW_MINIMIZE)
EndFunc

Func CloseApplication()
    Exit
EndFunc
