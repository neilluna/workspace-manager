Global $title = "Workspace Manager"
Global $version = "1.3"

AutoItSetOption("MustDeclareVars", 1)

#Include <Constants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>

Opt("GUIOnEventMode", 1)
Opt("GUIEventOptions", 1)

; Set the initial action variables.

Global $size_width = "No change"
Global $size_height = "No change"
Global $move_horizontal = "No change"
Global $move_vertical = "No change"

; Create the main window and all its controls.

GUICreate($title, 240, 195)

GUICtrlCreateGroup("Size", 10, 10, 220, 70)

GUICtrlCreateLabel("Width", 20, 28)
Global $size_width_combo = GUICtrlCreateCombo("No change", 90, 25, 130, 21, $CBS_DROPDOWNLIST)
GUICtrlSetData($size_width_combo, "480|640|800|960|1024|1280|Match width of ...", "No change")

GUICtrlCreateLabel("Height", 20, 53)
Global $size_height_combo = GUICtrlCreateCombo("No change", 90, 50, 130, 21, $CBS_DROPDOWNLIST)
GUICtrlSetData($size_height_combo, "360|480|600|720|768|960|1024|Match height of ...", "No change")

GUICtrlCreateGroup("Move", 10, 85, 220, 70)

GUICtrlCreateLabel("Horizontal", 20, 103)
Global $move_horizontal_combo = GUICtrlCreateCombo("No change", 90, 100, 130, 21, $CBS_DROPDOWNLIST)
GUICtrlSetData($move_horizontal_combo, "Center|Left Edge|Right Edge|Left justify with ...|Right justify with ...|Stack to the left of ...|Stack to the right of ...", "No change")

GUICtrlCreateLabel("Vertical", 20, 128)
Global $move_vertical_combo = GUICtrlCreateCombo("No change", 90, 125, 130, 21, $CBS_DROPDOWNLIST)
GUICtrlSetData($move_vertical_combo, "Center|Top Edge|Bottom Edge|Top justify with ...|Bottom justify with ...|Stack above ...|Stack below ...", "No change")

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

Func WidthChanged()
    $size_width = GUICtrlRead($size_width_combo)
EndFunc

Func HeightChanged()
    $size_height = GUICtrlRead($size_height_combo)
EndFunc

Func MoveHorizontalChanged()
    $move_horizontal = GUICtrlRead($move_horizontal_combo)
EndFunc

Func MoveVerticalChanged()
    $move_vertical = GUICtrlRead($move_vertical_combo)
EndFunc

Func AskForReferenceWindow($description)
    Return AskForWindow($description, "Press Ok, then select the reference window.")
EndFunc

Func AskForTargetWindow($description)
    Return AskForWindow($description, "Press Ok, then select the window to size and/or move.")
EndFunc

Func AskForWindow($description, $action)
    Local $prompt = $action
    If $description <> "" Then
        $prompt = $description & @CRLF & $action
    EndIf
    Local $choice = MsgBox(1, $title, $prompt, 60)
    If ($choice < 0 or $choice == 2) Then
        Return 0
    EndIf
    Sleep(5000)
    Return 1
EndFunc

Func GetDesktopWorkspacePos()
    Local $dll_rect = DllStructCreate("long; long; long; long")
    Local $result = DllCall("User32.dll", "int", "SystemParametersInfo", "uint", 48, "uint", 0, "ptr", DllStructGetPtr($dll_rect), "uint", 0)
    If @error Or $result[0] = 0 Then
        Return 0
    EndIf
    Local $rect[4] = [DllStructGetData($dll_rect, 1), DllStructGetData($dll_rect, 2), DllStructGetData($dll_rect, 3), DllStructGetData($dll_rect, 4)]
    $rect[2] = $rect[2] - $rect[0]
    $rect[3] = $rect[3] - $rect[1]
    Return $rect
EndFunc

Func ApplyChanges()
    Local $desktop_workspace = GetDesktopWorkspacePos()
    If $desktop_workspace = 0 Then
        MsgBox(0, $title, "Error:" & @CRLF & "The desktop workspace dimensions could not be determined", 60)
        Return
    EndIf

    Local $width_reference_window
    If $size_width = "Match width of ..." Then
        If Not AskForReferenceWindow("Match the width of another window.") Then
            Return
        EndIf
        $width_reference_window = WinGetPos("")
    EndIf

    Local $height_reference_window
    If $size_height = "Match height of ..." Then
        If Not AskForReferenceWindow("Match the height of another window.") Then
            Return
        EndIf
        $height_reference_window = WinGetPos("")
    EndIf

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

    If Not AskForTargetWindow("") Then
        Return
    EndIf
    Local $target_window = WinGetPos("")

    If $size_width = "Match width of ..." Then
        $target_window[2] = $width_reference_window[2]
    ElseIf $size_width <> "No change" Then
        $target_window[2] = $size_width
    EndIf

    If $size_height = "Match height of ..." Then
        $target_window[3] = $height_reference_window[3]
    ElseIf $size_height <> "No change" Then
        $target_window[3] = $size_height
    EndIf

    Select
        Case $move_horizontal = "Center"
            $target_window[0] = $desktop_workspace[0] + ($desktop_workspace[2] - $target_window[2]) / 2
        Case $move_horizontal = "Left Edge"
            $target_window[0] = $desktop_workspace[0]
        Case $move_horizontal = "Right Edge"
            $target_window[0] = $desktop_workspace[0] + $desktop_workspace[2] - $target_window[2]
        Case $move_horizontal = "Left justify with ..."
            $target_window[0] = $horizontal_reference_window[0]
        Case $move_horizontal = "Right justify with ..."
            $target_window[0] = $horizontal_reference_window[0] + $horizontal_reference_window[2] - $target_window[2]
        Case $move_horizontal = "Stack to the left of ..."
            $target_window[0] = $horizontal_reference_window[0] - $target_window[2]
        Case $move_horizontal = "Stack to the right of ..."
            $target_window[0] = $horizontal_reference_window[0] + $horizontal_reference_window[2]
    EndSelect

    Select
        Case $move_vertical = "Center"
            $target_window[1] = $desktop_workspace[1] + ($desktop_workspace[3] - $target_window[3]) / 2
        Case $move_vertical = "Top Edge"
            $target_window[1] = $desktop_workspace[1]
        Case $move_vertical = "Bottom Edge"
            $target_window[1] = $desktop_workspace[1] + $desktop_workspace[3] - $target_window[3]
        Case $move_vertical = "Top justify with ..."
            $target_window[1] = $vertical_reference_window[1]
        Case $move_vertical = "Bottom justify with ..."
            $target_window[1] = $vertical_reference_window[1] + $vertical_reference_window[3] - $target_window[3]
        Case $move_vertical = "Stack above ..."
            $target_window[1] = $vertical_reference_window[1] - $target_window[3]
        Case $move_vertical = "Stack below ..."
            $target_window[1] = $vertical_reference_window[1] + $vertical_reference_window[3]
    EndSelect

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
