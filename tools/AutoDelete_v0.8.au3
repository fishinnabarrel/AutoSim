; ------------------------------------------------------------------------
;
; Title:          AutoDelete.au3
; Developers:     David Pyke
; Date:           6/14/2007
; Last Update:    1/27/2011
; Version:        0.8.0
;
; Script Function: Automate player deletion.
;                  Delete players below AB, PA, IP or BF threshold.
;                  Updated to operate with DMB 10.
; ------------------------------------------------------------------------

; Check if script is already running.
_Singleton(@ScriptName, 0)

;~ #RequireAdmin

#include <Array.au3>
#include <Misc.au3>
#include <GUIConstantsEx.au3>

Opt("TrayIconDebug", 1)
Opt("MustDeclareVars", 1)   ; 0=no, 1=require pre-declare
Opt("WinTextMatchMode", 1)  ; 1=complete, 2=quick
Opt("WinTitleMatchMode", 1) ; 1=start, 2=subStr, 3=exact, 4=advanced

Global Const $Title = "AutoDelete"
Global Const $Version = "0.8"
Global Const $DMBTitle = "Diamond Mind Baseball"
Global Enum $PLAYERID, $ROLE, $IMAGETYPE, $FIRSTNAME, $LASTNAME, $TEAM, $POS, $AB, $PA, $IP, $BF
Global Const $NumProfileFields = 11
Global $g_deletedPlayerLog = ""
Global $g_playersDeleted = 0

HotKeySet("^!x", "ExitEarly")  ; Ctrl-Alt-x to exit script

; Check if DMB is running.
If Not WinExists($DMBTitle) Then
    MsgBox(16, $Title, "Diamond Mind Baseball is not running!" & @CRLF & @CRLF & _
        "Start DMB and try again.")
    Exit(1)
EndIf

#Region --- GUI related code start ---
Local $font = "Tahoma"
Local $fontBold = "Tahoma Bold"

Global $frmMain = GUICreate($Title & " v" & $Version, 314, 350, -1, -1)
GUISetFont(9, 400, 0, $font)

Global $mnuFileMenu = GuiCtrlCreateMenu("&File")
Global $mnuFileSeperator = GUICtrlCreateMenuitem ("",$mnuFileMenu, 1)
Global $mnuFileExitItem = GUICtrlCreateMenuitem ("E&xit",$mnuFileMenu)

Global $mnuOptionsMenu = GUICtrlCreateMenu("&Options")
Global $mnuOptionsExcludeCatchersItem = GUICtrlCreateMenuItem("&Exclude catchers", $mnuOptionsMenu, 1)
GUICtrlSetState($mnuOptionsExcludeCatchersItem, $GUI_UNCHECKED)
Global $mnuOptionsOnRosterItem = GUICtrlCreateMenuItem("Delete players on a &roster", $mnuOptionsMenu, 1)
GUICtrlSetState($mnuOptionsOnRosterItem, $GUI_UNCHECKED)

Global $mnuHelpMenu = GuiCtrlCreateMenu("&Help")
;~ Global $mnuHelpHelpItem = GUICtrlCreateMenuitem ("Help   F1",$mnuHelpMenu)
;~ Global $mnuHelpSeperator = GUICtrlCreateMenuitem ("",$mnuHelpMenu, 1)
Global $mnuHelpAboutItem = GUICtrlCreateMenuitem ("&About AutoDelete",$mnuHelpMenu)

Global $lblHeader = GUICtrlCreateLabel("Choose Options:", 12, 12, 300, 22)
GUICtrlSetFont(-1, 12, 400, 0, $fontBold)

Global $grpBatters = GUICtrlCreateGroup("", 12, 36, 290, 90)
Global $chkBatters = GUICtrlCreateCheckbox("Delete Batters", 22, 59, 150, 17)
GUICtrlSetFont(-1, 9, 400, 0, $fontBold)
Global $lblBatterPA = GUICtrlCreateLabel("Min:", 45, 86, 25, 22)
Global $txtBatterPA = GUICtrlCreateInput("75", 80, 84, 40, 22)
Global $rdoBatterPA = GUICtrlCreateRadio("PA", 130, 84, 40, 25)
Global $rdoBatterAB = GUICtrlCreateRadio("AB", 177, 84, 40, 25)

Global $grpPitchers = GUICtrlCreateGroup("", 12, 130, 290, 90)
Global $chkPitchers = GUICtrlCreateCheckbox("Delete Pitchers", 22, 153, 150, 17)
GUICtrlSetFont(-1, 9, 400, 0, $fontBold)
Global $lblPitcherBF = GUICtrlCreateLabel("Min:", 45, 180, 25, 22)
Global $txtPitcherBF = GUICtrlCreateInput("75", 80, 178, 40, 22)
Global $rdoPitcherBF = GUICtrlCreateRadio("BF", 130, 178, 40, 25)
Global $rdoPitcherIP = GUICtrlCreateRadio("IP", 177, 178, 40, 25)

Global $chkMultiteam = GUICtrlCreateCheckbox("Delete Multi-team players", 22, 240, 300, 17)
GUICtrlSetFont(-1, 9, 400, 0, $fontBold)

GUICtrlCreateGraphic(12, 272)
GUICtrlSetGraphic(-1, $GUI_GR_LINE, 290, 0)

Global $btnRun = GUICtrlCreateButton("Run", 140, 285, 75, 25)
Global $btnClose = GUICtrlCreateButton("Close", 225, 285, 75, 25)

GUICtrlSetState($rdoBatterPA, $GUI_CHECKED)
GUICtrlSetState($rdoPitcherBF, $GUI_CHECKED)
GUICtrlSetState($rdoBatterPA, $GUI_DISABLE)
GUICtrlSetState($rdoBatterAB, $GUI_DISABLE)
GUICtrlSetState($txtBatterPA, $GUI_DISABLE)
GUICtrlSetState($rdoPitcherBF, $GUI_DISABLE)
GUICtrlSetState($rdoPitcherIP, $GUI_DISABLE)
GUICtrlSetState($txtPitcherBF, $GUI_DISABLE)

GUISetState(@SW_SHOW)
#EndRegion --- GUI related code end ---

Local $guiMsg = 0
While True
    $guiMsg = GUIGetMsg()
    Select
        Case $guiMsg = $GUI_EVENT_CLOSE
            ExitLoop

        Case $guiMsg = $mnuFileExitItem Or $guimsg = $btnClose
            ExitLoop

        Case $guiMsg = $chkBatters
            If GUICtrlRead($chkBatters) = $GUI_CHECKED Then
                GUICtrlSetState($rdoBatterPA, $GUI_ENABLE)
                GUICtrlSetState($rdoBatterAB, $GUI_ENABLE)
                GUICtrlSetState($txtBatterPA, $GUI_ENABLE)
            Else
                GUICtrlSetState($rdoBatterPA, $GUI_DISABLE)
                GUICtrlSetState($rdoBatterAB, $GUI_DISABLE)
;~              GUICtrlSetData($txtBatterPA, "0")
                GUICtrlSetState($txtBatterPA, $GUI_DISABLE)
            EndIf

        Case $guiMsg = $chkPitchers
            If GUICtrlRead($chkPitchers) = $GUI_CHECKED Then
                GUICtrlSetState($rdoPitcherBF, $GUI_ENABLE)
                GUICtrlSetState($rdoPitcherIP, $GUI_ENABLE)
                GUICtrlSetState($txtPitcherBF, $GUI_ENABLE)
            Else
                GUICtrlSetState($rdoPitcherBF, $GUI_DISABLE)
                GUICtrlSetState($rdoPitcherIP, $GUI_DISABLE)
;~              GUICtrlSetData($txtPitcherBF, "0")
                GUICtrlSetState($txtPitcherBF, $GUI_DISABLE)
            EndIf

        Case $guiMsg = $btnRun
            If GUICtrlRead($chkBatters) = $GUI_UNCHECKED And _
                GUICtrlRead($chkPitchers) = $GUI_UNCHECKED And _
                GUICtrlRead($chkMultiteam) = $GUI_UNCHECKED Then
                MsgBox(48, $Title, "Choose something to delete.")
            Else
                ProcessPlayers()
            EndIf
            GUISetState(@SW_RESTORE, $frmMain)

        Case $guiMsg = $mnuOptionsOnRosterItem
            If BitAnd(GUICtrlRead($mnuOptionsOnRosterItem), $GUI_CHECKED) = $GUI_CHECKED Then
                GUICtrlSetState($mnuOptionsOnRosterItem, $GUI_UNCHECKED)
            Else
                GUICtrlSetState($mnuOptionsOnRosterItem, $GUI_CHECKED)
            EndIf

        Case $guiMsg = $mnuOptionsExcludeCatchersItem
            If BitAnd(GUICtrlRead($mnuOptionsExcludeCatchersItem), $GUI_CHECKED) = $GUI_CHECKED Then
                GUICtrlSetState($mnuOptionsExcludeCatchersItem, $GUI_UNCHECKED)
            Else
                GUICtrlSetState($mnuOptionsExcludeCatchersItem, $GUI_CHECKED)
            EndIf

        Case $guiMsg = $mnuHelpAboutItem
            DisplayAbout()

    EndSelect
WEnd

Exit(0)

; ----------------------------------------------------------------------------
; functions below

Func ProcessPlayers()
    ; Goto DMB's Organizer window and process the list of players.
    ; Returns: None

    Local $minBatter = Int(GUICtrlRead($txtBatterPA))
    Local $minPitcher = Int(GUICtrlRead($txtPitcherBF))
    Local $playersToDelete = 0
    Local $deleteBatters = False
    Local $deletePitchers = False
    Local $deleteMulti = False
    Local $deleteOnRoster = False
    Local $excludeCatchers = False
    Local $usePA = True
    Local $useBF = True

    If GUICtrlRead($chkBatters) = $GUI_CHECKED Then $deleteBatters = True
    If GUICtrlRead($chkPitchers) = $GUI_CHECKED Then $deletePitchers = True
    If GUICtrlRead($chkMultiteam) = $GUI_CHECKED Then $deleteMulti = True
    If BitAnd(GUICtrlRead($mnuOptionsOnRosterItem), $GUI_CHECKED) = $GUI_CHECKED Then $deleteOnRoster = True
    If BitAnd(GUICtrlRead($mnuOptionsExcludeCatchersItem), $GUI_CHECKED) = $GUI_CHECKED Then $excludeCatchers = True

    If GUICtrlRead($rdoBatterAB) = $GUI_CHECKED Then $usePA = False
    If GUICtrlRead($rdoPitcherIP) = $GUI_CHECKED Then $useBF = False

    OpenPlayer()

    Local $previousLocalID = -99
    Local $playerProfile[$NumProfileFields]

    ExaminePlayer($playerProfile, $deleteBatters, $deletePitchers, $deleteOnRoster)
    While $playerProfile[$PLAYERID] <> $previousLocalID
        $previousLocalID = $playerProfile[$PLAYERID]
        If ($deleteOnRoster = False) And ($playerProfile[$TEAM] <> "") Then
            ; This player is on a roster and will not be deleted.
            Send("{DOWN}")
        ElseIf ($excludeCatchers = True) And ($playerProfile[$POS] = "c") Then
            ; Catchers excluded, player will not be deleted.
            Send("{DOWN}")
        Else
            ; For 'Dual' role players assign role based on primary position
            If ($playerProfile[$ROLE] = "Dual") Then
                If ($playerProfile[$POS] = "sp" OR $playerProfile[$POS] = "rp" OR _
                    $playerProfile[$POS] = "mr" OR $playerProfile[$POS] = "cl") Then
                    $playerProfile[$ROLE] = "Pitcher"
                Else
                    $playerProfile[$ROLE] = "Batter"
                EndIf
            EndIf

            If ($deleteMulti = True) And ($playerProfile[$IMAGETYPE] = "Traded") Then
                ; Delete single-team record for multi-team players
                DeletePlayer()
                LogPlayer($playerProfile)
            ElseIf ($deleteBatters = True) And ($playerProfile[$ROLE] = "Batter") Then
                If $usePA = True Then
                    If ($minBatter > 0) And ($playerProfile[$PA] < $minBatter) Then
;~                      _DebugDisplay($playerProfile)
                        DeletePlayer()
                        LogPlayer($playerProfile)
                    Else
                        Send("{DOWN}")
                    EndIf
                Else
                    If ($minBatter > 0) And ($playerProfile[$AB] < $minBatter) Then
                        DeletePlayer()
                        LogPlayer($playerProfile)
                    Else
                        Send("{DOWN}")
                    EndIf
                EndIf
            ElseIf ($deletePitchers = True) And ($playerProfile[$ROLE] = "Pitcher") Then
                If $useBF = True Then
                    If ($minPitcher > 0) And ($playerProfile[$BF] < $minPitcher) Then
;~                      _DebugDisplay($playerProfile)
                        DeletePlayer()
                        LogPlayer($playerProfile)
                    Else
                        Send("{DOWN}")
                    EndIf
                Else
                    If ($minPitcher > 0) And ($playerProfile[$IP] < $minPitcher) Then
                        DeletePlayer()
                        LogPlayer($playerProfile)
                    Else
                        Send("{DOWN}")
                    EndIf
                EndIf
            Else
                Send("{DOWN}")
            EndIf
        EndIf

        ExaminePlayer($playerProfile, $deleteBatters, $deletePitchers, $deleteOnRoster)

    WEnd

    Savelog()

EndFunc

Func OpenPlayer()
    ; Open player window.
    ; Returns: None
    If ControlCommand($DMBTitle, "Organizer", "", "IsVisible") Then
        WinActivate($DMBTitle, "Organizer")
        WinWaitActive($DMBTitle, "Organizer")
        ControlClick($DMBTitle, "Organizer", "[ID:59905]", "left", 1, 175, 8)
    Else
        WinActivate($DMBTitle, "")
        WinWaitActive($DMBTitle, "")
;~      WinSetState($DMBTitle, "", @SW_MAXIMIZE)
        WinMenuSelectItem($DMBTitle, "","&View", "&Organizer...")
        WinWait($DMBTitle, "Organizer")
;~      WinMenuSelectItem($DMBTitle, "","&Window", "&Tile")
        ControlClick($DMBTitle, "Organizer", "[ID:59905]", "left", 1, 175, 8)
    EndIf
    Return
EndFunc

Func ExaminePlayer(ByRef $playerProfile, $deleteBatters, $deletePitchers, $deleteOnRoster)
    ; Determines if player is batter or pitcher.
    ; Reads Real-life PAs and IPs.
    ; Returns: ByRef array $playerProfile
    WinWait($DMBTitle, "Organizer")
    ControlClick($DMBTitle, "Organizer", "Modify")
    Send("g")
    WinWait("Modify Player")
    $playerProfile[$PLAYERID] = ControlGetText("Modify Player", "", "[ID:1380]")
    $playerProfile[$ROLE] = StringStripWS(ControlGetText("Modify Player", "", "[ID:1375]"), 3)
    $playerProfile[$IMAGETYPE] = StringStripWS(ControlGetText("Modify Player", "", "[ID:2759]"), 3)
    $playerProfile[$FIRSTNAME] = StringStripWS(ControlGetText("Modify Player", "", "[ID:1372]"), 3)
    $playerProfile[$LASTNAME] = StringStripWS(ControlGetText("Modify Player", "", "[ID:1373]"), 3)
    $playerProfile[$TEAM] = StringStripWS(ControlGetText("Modify Player", "", "[ID:1387]"), 3)
    $playerProfile[$POS] = StringStripWS(ControlGetText("Modify Player", "", "[ID:1379]"), 3)

    ControlClick("Modify Player", "", "OK")
    WinWait($DMBTitle, "Organizer")

    If $deleteOnRoster = False And $playerProfile[$TEAM] <> "" Then
        ; No need to examine stats.
        ; This player is on a roster and will not be deleted.
    Else
        $playerProfile[$IP] = ""
        $playerProfile[$BF] = ""
        $playerProfile[$AB] = ""
        $playerProfile[$PA] = ""

        If ($playerProfile[$ROLE] <> "Pitcher") And ($deleteBatters = True) Then
            ControlClick($DMBTitle, "Organizer", "Modify")
            Send("r")
            WinWait("Modify real-life statistics", "Batting Statistics")
            $playerProfile[$AB] = ControlGetText("Modify real-life statistics", "Batting Statistics", "[ID:1200]")
            ; PA = AB + BB + HBP + SH + SF + CI
            $playerProfile[$PA] = $playerProfile[$AB] + _
                ControlGetText("Modify real-life statistics", "Batting Statistics", "[ID:1244]") + _
                ControlGetText("Modify real-life statistics", "Batting Statistics", "[ID:1250]") + _
                ControlGetText("Modify real-life statistics", "Batting Statistics", "[ID:1280]") + _
                ControlGetText("Modify real-life statistics", "Batting Statistics", "[ID:1285]") + _
                ControlGetText("Modify real-life statistics", "Batting Statistics", "[ID:1290]")
            ControlClick("Modify real-life statistics", "", "OK")
        EndIf
        If ($playerProfile[$ROLE] <> "Batter") And ($deletePitchers = True) Then
            ControlClick($DMBTitle, "Organizer", "Modify")
            Send("r")
            WinWait("Modify real-life statistics")
            ControlClick("Modify real-life statistics", "", "[ID:12320]", "left", 1, 400, 33)
            WinWait("Modify real-life statistics", "Pitching Statistics")
            $playerProfile[$IP] = ControlGetText("Modify real-life statistics", "Pitching Statistics", "[ID:1275]")
            $playerProfile[$BF] = ControlGetText("Modify real-life statistics", "Pitching Statistics", "[ID:1269]")
            ControlClick("Modify real-life statistics", "", "OK")
        EndIf

        WinWait($DMBTitle, "Organizer")
;~      _DebugDisplay($playerProfile)
    EndIf

    Return
EndFunc

Func DeletePlayer()
    ; Deletes currently highlighted player.
    ; Returns: None
    WinWait($DMBTitle, "Organizer")
    ControlClick($DMBTitle, "Organizer", "Delete")
    WinWait("Baseball", "This deletion cannot be reversed")
    ControlClick("Baseball", "This deletion cannot be reversed", "&Yes")
    WinWait($DMBTitle, "Organizer")
    Return
EndFunc

Func LogPlayer(Const $playerProfile)
    ; Adds deleted player to log file and increments count.
    ; Return: None
    Local $text = '"' & $playerProfile[$PLAYERID] & '","' & _
        $playerProfile[$IMAGETYPE] & '","' & _
        $playerProfile[$ROLE] & '","' & _
        $playerProfile[$FIRSTNAME] & " " & _
        $playerProfile[$LASTNAME] & '","' & _
        $playerProfile[$TEAM] & '","' & _
        $playerProfile[$POS] & '","' & _
        $playerProfile[$AB] & '","' & _
        $playerProfile[$PA] & '","' & _
        $playerProfile[$IP] & '","' & _
        $playerProfile[$BF] & '"' & @CRLF

    $g_deletedPlayerLog &= $text
    $g_playersDeleted += 1
    Return
EndFunc

Func SaveLog()
    ; Write deleted player log to file.
    Local $header = '"PlayerID","ImageType","Role","Name","Team","Pos","AB","PA","IP","BF"' & @CRLF
    MsgBox(0, $Title & " v" & $Version, $g_playersDeleted & " players deleted.")
    Local $file = FileOpen("DeletedPlayerLog.csv", 2)
    If $file = -1 Then
        MsgBox(0, "Error", "Unable to create log file.")
        Exit(1)
    EndIf
    FileWrite($file, $header)
    FileWrite($file, $g_deletedPlayerLog)
    FileClose($file)
    Return
EndFunc

Func DisplayAbout()
    ; Displays an About message.
    ; Return value: None
    Local $reportText = $Title & @CRLF & _
        "Version: " & $Version & @CRLF & _
        "August 2011" & @CRLF & @CRLF & _
        "David Pyke"
    MsgBox(0, "About " & $Title, $reportText)
    Return 0
EndFunc

Func DisplayHelp()
    ; Displays a help message.
    ; Return value: None
    Return
EndFunc

Func ExitEarly()
    ; Exists script early - from hotkey.
    MsgBox(0, $Title, "Program stopped before completion.")
    Savelog()
    Exit(1)
EndFunc

Func _DebugDisplay($text)
    ; Display debugging info
    If IsArray($text) Then
        _ArrayDisplay($text, "Debugging")
    Else
        MsgBox(64+262144, "Debugging", $text)
    Endif
    Return
EndFunc
