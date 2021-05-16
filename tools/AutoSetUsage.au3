#cs
-----------------------------------------------------------------------------
Title:            AutoSet.au3
Developer:        David Pyke
Date:             September 10, 2008
Modified:         November 16, 2008
Version:          0.6.0

Description: Automate bulk settings changes.
             Manager Profile - Batter and Pitcher usage modes.
-----------------------------------------------------------------------------
Copyright (C) 2021 David Pyke.

This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
-----------------------------------------------------------------------------
#ce

#include <Misc.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

;~ Opt("TrayIconDebug", 1)

Opt("MustDeclareVars", 1)   ; 0=no, 1=require pre-declare
Opt("WinTextMatchMode", 1)  ; 1=complete, 2=quick
Opt("WinTitleMatchMode", 1) ; 1=start, 2=subStr, 3=exact

Global Const $Title = "AutoSet"
Global Const $Version = "0.6"
Global Const $DMBTitle = "Diamond Mind Baseball"

; Check if script is already running.
_Singleton(@ScriptName, 0)

HotKeySet("^!x", "ExitEarly")  ; Ctrl-Alt-x to exit script

; Check if DMB is running.
If Not WinExists($DMBTitle) Then
    MsgBox(16, $Title, "Diamond Mind Baseball 9 is not running!" & @CRLF & @CRLF & _
        "Start DMB9 and try again.")
    Exit(1)
EndIf

#region --- GUI related code start ---
Global $frmMain = GUICreate($Title & " v" & $Version, 325, 295, -1, -1)

Global $mnuFileMenu = GuiCtrlCreateMenu("&File")
Global $mnuFileSeperator = GUICtrlCreateMenuitem ("",$mnuFileMenu, 1)
Global $mnuFileExitItem = GUICtrlCreateMenuitem ("E&xit",$mnuFileMenu)
Global $mnuHelpMenu = GuiCtrlCreateMenu("&Help")
;~ Global $mnuHelpHelpItem = GUICtrlCreateMenuitem ("Help   F1",$mnuHelpMenu)
;~ Global $mnuHelpSeperator = GUICtrlCreateMenuitem ("",$mnuHelpMenu, 1)
Global $mnuHelpAboutItem = GUICtrlCreateMenuitem ("&About AutoSet",$mnuHelpMenu)

Global $lblTitle = GUICtrlCreateLabel("Global Settings:", 15, 15, 100, 25)

Global $rdoSortPlayers = GUICtrlCreateRadio("Arrange Players:", 15, 40, 150, 25)
Global $rdoStatusCodes = GUICtrlCreateRadio("Copy Stauts Codes:", 15, 75, 150, 25)
Global $rdoBatting = GUICtrlCreateRadio("Batting Usage mode:", 15, 110, 150, 25)
Global $rdoPitching = GUICtrlCreateRadio("Pitching Usage mode:", 15, 145, 150, 25)
GUICtrlSetState($rdoSortPlayers, $GUI_CHECKED)
GUICtrlSetState($rdoStatusCodes, $GUI_DISABLE)

Global $cboSortPlayers = GUICtrlCreateCombo("", 170, 40, 140, 25)
GUICtrlSetData(-1, "Alphabetically|By role|By primary position", "Alphabetically")
GUICtrlSetState($cboSortPlayers, $GUI_ENABLE)

Global $cboStatusCodes = GUICtrlCreateCombo("", 170, 75, 140, 25)
GUICtrlSetData(-1, "To pre-season|To current|", "")
GUICtrlSetState($cboStatusCodes, $GUI_DISABLE)

Global $cboBatting = GUICtrlCreateCombo("", 170, 110, 140, 25)
GUICtrlSetData(-1, "Track starts|Game by game", "Game by game")
GUICtrlSetState($cboBatting, $GUI_DISABLE)

Global $cboPitching = GUICtrlCreateCombo("", 170, 145, 140, 25)
GUICtrlSetData(-1, "Strict|Skip|Time", "Strict")
GUICtrlSetState($cboPitching, $GUI_DISABLE)

Global $lblRotation = GUICtrlCreateLabel("Rotation size:", 35, 180, 100, 25)
Global $cboRotation = GUICtrlCreateCombo("", 260, 175, 50, 23)
GUICtrlSetData(-1, "1|2|3|4|5", "5")
GUICtrlSetState($lblRotation, $GUI_DISABLE)
GUICtrlSetState($cboRotation, $GUI_DISABLE)

Global $btnStart = GUICtrlCreateButton("Start", 100, 230, 100, 28)
Global $btnExit = GUICtrlCreateButton("Exit", 210, 230, 100, 28)

GUISetState(@SW_SHOW)
#endregion --- GUI related code end ---

Local $teamCount = 0
Local $selection = 0
Local $guiMsg = 0
While True
    $guiMsg = GUIGetMsg()
    Select
        Case $guiMsg = $btnStart
            Select
                Case GUICtrlRead($rdoSortPlayers) = $GUI_CHECKED
                    Local $sortMethod = GUICtrlRead($cboSortPlayers)
                    $selection = MsgBox(32 + 1, $Title & " v" & $Version, _
                            "All teams will be sorted: " & $sortMethod)
                    If $selection = 1 Then $teamCount = SortPlayers($sortMethod)

                Case GUICtrlRead($rdoStatusCodes) = $GUI_CHECKED

                Case GUICtrlRead($rdoBatting) = $GUI_CHECKED
                    Local $batterUsageMode = GUICtrlRead($cboBatting)
                    $selection = MsgBox(32 + 1, $Title & " v" & $Version, _
                            "All teams will be set to: " & $batterUsageMode)
                    If $selection = 1 Then $teamCount = SetBatterUsage($batterUsageMode)

                Case GUICtrlRead($rdoPitching) = $GUI_CHECKED
                    Local $pitcherUsageMode = GUICtrlRead($cboPitching)
                    Local $rotationSize = GUICtrlRead($cboRotation)
                    $selection = MsgBox(32 + 1, $Title & " v" & $Version, _
                            "All teams will be set to: " & $pitcherUsageMode & _
                            " (" & $rotationSize & ")")
                    If $selection = 1 Then $teamCount = _
                            SetPitcherUsage($pitcherUsageMode, $rotationSize)

            EndSelect
            MsgBox(64, $Title & " v" & $Version, "Teams updated: " & $teamCount)
            GUISetState(@SW_RESTORE, $frmMain)

        Case $guiMsg = $rdoSortPlayers And BitAND(GUICtrlRead($rdoSortPlayers), $GUI_CHECKED) = $GUI_CHECKED
            GUICtrlSetState($cboSortPlayers, $GUI_ENABLE)
            GUICtrlSetState($cboStatusCodes, $GUI_DISABLE)
            GUICtrlSetState($cboBatting, $GUI_DISABLE)
            GUICtrlSetState($cboPitching, $GUI_DISABLE)
            GUICtrlSetState($lblRotation, $GUI_DISABLE)
            GUICtrlSetState($cboRotation, $GUI_DISABLE)

        Case $guiMsg = $rdoStatusCodes And BitAND(GUICtrlRead($rdoStatusCodes), $GUI_CHECKED) = $GUI_CHECKED
            GUICtrlSetState($cboSortPlayers, $GUI_DISABLE)
            GUICtrlSetState($cboStatusCodes, $GUI_ENABLE)
            GUICtrlSetState($cboBatting, $GUI_DISABLE)
            GUICtrlSetState($cboPitching, $GUI_DISABLE)
            GUICtrlSetState($lblRotation, $GUI_DISABLE)
            GUICtrlSetState($cboRotation, $GUI_DISABLE)

        Case $guiMsg = $rdoBatting And BitAND(GUICtrlRead($rdoBatting), $GUI_CHECKED) = $GUI_CHECKED
            GUICtrlSetState($cboSortPlayers, $GUI_DISABLE)
            GUICtrlSetState($cboStatusCodes, $GUI_DISABLE)
            GUICtrlSetState($cboBatting, $GUI_ENABLE)
            GUICtrlSetState($cboPitching, $GUI_DISABLE)
            GUICtrlSetState($lblRotation, $GUI_DISABLE)
            GUICtrlSetState($cboRotation, $GUI_DISABLE)

        Case $guiMsg = $rdoPitching And BitAND(GUICtrlRead($rdoPitching), $GUI_CHECKED) = $GUI_CHECKED
            GUICtrlSetState($cboSortPlayers, $GUI_DISABLE)
            GUICtrlSetState($cboStatusCodes, $GUI_DISABLE)
            GUICtrlSetState($cboBatting, $GUI_DISABLE)
            GUICtrlSetState($cboPitching, $GUI_ENABLE)
            GUICtrlSetState($lblRotation, $GUI_ENABLE)
            GUICtrlSetState($cboRotation, $GUI_ENABLE)

        Case $guiMsg = $mnuHelpAboutItem
            DisplayAbout()

        Case $guiMsg = $GUI_EVENT_CLOSE Or $guiMsg = $mnuFileExitItem Or $guiMsg = $btnExit
            ExitLoop

    EndSelect
WEnd

Exit(0)


; ----------------------------------------------------------------------------
; functions below

Func SetBatterUsage($batterUsageMode)
    ; Open or activate Roster/Manager profile screen and set batting usage mode.
    ; Returns: Number of teams updated

    Local Const $DepthChartTabCoord = 240
    OpenRosterWindow($DepthChartTabCoord)

    Local $startTeam = ControlGetText($DMBTitle, "Team roster - ", "[ID:1904]")
;~  MsgBox(64, $Title & " v" & $Version, "First team: " & $startTeam)

    Local $teamCount = 0
    Local $currentTeam = ""
    While $currentTeam <> $startTeam

        ; Set the usage mode
        ControlCommand($DMBTitle, "Team roster - ", "[ID:1991]", "SelectString", $batterUsageMode)
        Sleep(100)

        ; Goto next team
        ControlClick($DMBTitle, "Team roster - ", "[ID:1905]")
        Sleep(100)
        If WinExists("Baseball", "You have made changes") Then
            ControlClick("Baseball", "You have made changes", "&Yes")
            WinWaitActive($DMBTitle, "Team roster - ")
        EndIf
        $currentTeam = ControlGetText($DMBTitle, "Team roster - ", "[ID:1904]")
        $teamCount += 1
    WEnd
    Return $teamCount
EndFunc

Func SetPitcherUsage($pitcherUsageMode, $rotationSize)
    ; Open or activate Roster/Manager profile screen and set pitching usage mode.
    ; Returns: Number of teams updated

    Local Const $PitchingChartTabCoord = 75
    OpenRosterWindow($PitchingChartTabCoord)

    Local $startTeam = ControlGetText($DMBTitle, "Team roster - ", "[ID:1904]")
;~  MsgBox(64, $Title & " v" & $Version, "First team: " & $startTeam)

    Local $teamCount = 0
    Local $currentTeam = ""
    While $currentTeam <> $startTeam

        ; Set the usage mode
        ControlCommand($DMBTitle, "Team roster - ", "[ID:1959]", "SelectString", $pitcherUsageMode)
        Sleep(100)

        ; Set the rotation size
        ControlFocus($DMBTitle, "Team roster - ", "[ID:1960]")
        ControlSetText($DMBTitle, "Team roster - ", "[ID:1960]", $rotationSize)
        Sleep(100)

        ; Goto next team
        ControlClick($DMBTitle, "Team roster - ", "[ID:1905]")
        Sleep(100)
        If WinExists("Baseball", "You have made changes") Then
            ControlClick("Baseball", "You have made changes", "&Yes")
            WinWaitActive($DMBTitle, "Team roster - ")
        EndIf
        $currentTeam = ControlGetText($DMBTitle, "Team roster - ", "[ID:1904]")
        $teamCount += 1
    WEnd
    Return $teamCount
EndFunc

Func SortPlayers($sortMethod)
    ; Open or activate Roster/Manager profile screen and sort player records.
    ; Returns: Number of teams updated

    Local Const $RosterTabCoord = 10
    OpenRosterWindow($RosterTabCoord)

    Local $startTeam = ControlGetText($DMBTitle, "Team roster - ", "[ID:1904]")
;~  MsgBox(64, $Title & " v" & $Version, "First team: " & $startTeam)

    Local $teamCount = 0
    Local $currentTeam = ""
    While $currentTeam <> $startTeam

        ; Arrange players
        ControlClick($DMBTitle, "Team roster - ", "[ID:2520]", "right")
        Send("n")
        Select
            Case $sortMethod = "Alphabetically"
                Send("a")

            Case $sortMethod = "By role"
                Send("r")

            Case $sortMethod = "By primary position"
                Send("p")

        EndSelect
        Sleep(100)

        ; Goto next team
        ControlClick($DMBTitle, "Team roster - ", "[ID:1905]")
        WinWaitActive($DMBTitle, "Team roster - ")
        $currentTeam = ControlGetText($DMBTitle, "Team roster - ", "[ID:1904]")
        $teamCount += 1
    WEnd
    Return $teamCount
EndFunc

Func OpenRosterWindow($tabCoord)
    ; Open Roster/Manager profile screen.
    ; Returns: None

    ; Is Roster/manager profile window open?
    If ControlCommand($DMBTitle, "Team roster - ", "", "IsVisible") Then
        WinActivate($DMBTitle, "Team roster - ")
        WinWaitActive($DMBTitle, "Team roster - ")
        ; Click on tab
        ControlClick($DMBTitle, "Team roster - ", "[ID:59952]", "left", 1, $tabCoord, 12)
    Else
        WinActivate($DMBTitle, "")
        WinWaitActive($DMBTitle, "")
;~      WinSetState($DMBTitle, "", @SW_MAXIMIZE)
        WinMenuSelectItem($DMBTitle, "","&View", "&Roster / manager profile")
        WinWait("Roster/Manager Profile")
        ControlClick("Roster/Manager Profile", "", "OK")
        WinWait($DMBTitle, "Team roster - ")
;~      WinMenuSelectItem($DMBTitle, "","&Window", "&Tile")
        ; Click on tab
        ControlClick($DMBTitle, "Team roster - ", "[ID:59952]", "left", 1, $tabCoord, 12)
    EndIf
    Return
EndFunc

Func DisplayAbout()
    ; Displays an About message.
    ; Return value: None
    Local $reportText = $Title & @CRLF & _
        "Version: " & $Version & @CRLF & _
        "November 2008" & @CRLF & @CRLF & _
        "David Pyke"
    MsgBox(0, "About " & $Title, $reportText)
    Return 0
EndFunc

Func DisplayHelp()
    ; Displays a help message.
    ; Return vale: None
    Return
EndFunc

Func ExitEarly()
    ; Exists script early - from hotkey.
    MsgBox(0, $Title, "Program stopped before completion.")
    Exit(1)
EndFunc

Func _DebugDisplay($text)
    ; Display debugging info
    MsgBox(64+262144, "Debugging", $text)
    Return
EndFunc

