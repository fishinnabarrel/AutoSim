#cs
-------------------------------------------------------------------------------
 Title:         AutoSim.au3
 Developer:     David Pyke
 Created:       February 26, 2006
 Last Update:   December 8, 2021
 Version:       12.0

 Description: Automate multiple season simulations with Diamond Mind
              Baseball and import into Diamond Mind Baseball
              Encyclopedia. GUI version.

May 16, 2021 (Version 12.0):
- Update for DMB version 12
-------------------------------------------------------------------------------
Copyright (C) 2022 David Pyke.

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
-------------------------------------------------------------------------------
#ce

#include <GuiConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <Date.au3>
#include <Misc.au3>

Opt( "GUIOnEventMode", 1 )      ; 0=disable, 1=enable
Opt( "MustDeclareVars", 1 )     ; 0=no, 1=require pre-declare
Opt( "WinTextMatchMode", 2 )    ; 1=complete, 2=quick
Opt( "WinTitleMatchMode", 2 )   ; 1=start, 2=subStr, 3=exact, 4=advanced
Opt( "TrayAutoPause", 0 )       ; 0=no pause, 1=pause (default)
Opt( "TrayOnEventMode", 1 )     ; 0=disable (default), 1=enable
Opt( "TrayMenuMode", 1 )        ; ;0=append, 1=no default menu
;~ Opt( "WinWaitDelay", 500 )       ; 500 milliseconds

Global Const $SCRIPT_TITLE = "AutoSim"
Global Const $SCRIPT_VERSION = "12.0"
Global Const $DMB_TITLE = "Diamond Mind Baseball"
Global Const $DMBENC_TITLE = "Diamond Mind Baseball Encyclopedia"
Global Enum  $DMB_PATH = 1, $ENC_PATH, $SEASONS, $START, _
             $IS_RESTART, $LEFT, $RESTART, $GBGSTATS, $BOXSCORES
Global $g_startupSettings[10][2]
Global $g_dmbInstallPath = "C:\dmb12"
Global $g_dmbEncInstallPath = "C:\dmbenc12"
Global $g_saveOnExit = 0
Global $g_schedule = 0

; Check for instance of Autosim already running
_Singleton( @ScriptName, 0 )

HotKeySet( "^!x", "ExitClicked" )           ; Ctrl-Alt-x to exit script.

#region --- GUI related code start ---
Local $font = "Tahoma"
Local $fontBold = "Tahoma Bold"

Global $frmMain = GUICreate( $SCRIPT_TITLE & " " & $SCRIPT_VERSION, 487, 326, -1, -1 )
GUISetFont( 10, 400, 0, $font )

Global $mnuFileMenu = GuiCtrlCreateMenu( "&File" )
Global $mnuFileDMBPathItem = GUICtrlCreateMenuItem( "Update path to &DMB Folder", $mnuFileMenu )
Global $mnuFileENCPathItem = GUICtrlCreateMenuItem( "Update path to DMB&ENC Folder", $mnuFileMenu )
Global $mnuFileSeperator = GUICtrlCreateMenuitem ( "", $mnuFileMenu )
Global $mnuFileExitItem = GUICtrlCreateMenuitem( "E&xit" & @TAB & "Ctrl+Alt+X", $mnuFileMenu )

Global $mnuOptionsMenu = GUICtrlCreateMenu( "&Options" )
Global $mnuOptionsGBGStatsItem = GUICtrlCreateMenuItem( "Include &game-by-game stats", $mnuOptionsMenu )
Global $mnuOptionsBoxscoreItem = GUICtrlCreateMenuItem( "Include &boxscores", $mnuOptionsMenu )
Global $mnuOptionsRefreshItem = GUICtrlCreateMenuitem ( "&Refresh active database" & @TAB & "F5", $mnuOptionsMenu )

Global $mnuHelpMenu = GuiCtrlCreateMenu( "&Help" )
;~ Global $mnuHelpHelpItem = GUICtrlCreateMenuitem ("Help",$mnuHelpMenu)
;~ Global $mnuHelpSeperator = GUICtrlCreateMenuitem ("",$mnuHelpMenu, 1)
Global $mnuHelpAboutItem = GUICtrlCreateMenuitem( "&About AutoSim" & @TAB & "F1", $mnuHelpMenu )

Global $grpTop = GuiCtrlCreateGroup( "", 20, 5, 445, 80 )
Global $lblSeasons = GuiCtrlCreateLabel( "Number of seasons to simulate:", 30, 25, 365, 20 )
GUICtrlSetFont( -1, 10, 700, 0, $font )
Global $txtSeasons = GuiCtrlCreateInput( "10", 405, 25, 50, 22 )
GUICtrlSetState( -1, $GUI_FOCUS )
Global $lblStart = GuiCtrlCreateLabel( "Start numbering seasons at:", 30, 56, 365, 20 )
GUICtrlSetFont( -1, 10, 700, 0, $font )
Global $txtStart = GuiCtrlCreateInput( "1", 405, 55, 50, 22 )

Global $grpBottom = GuiCtrlCreateGroup( "", 20, 90, 445, 147 )
Global $lblDatabase = GuiCtrlCreateLabel( "Active DMB database: ", 30, 111, 365, 20 )
GUICtrlSetFont( -1, 10, 700, 0, $font )
Global $txtDBFrame = GUICtrlCreateLabel( "", 30, 136, 400, 28, $SS_ETCHEDFRAME )
Global $txtDatabase = GUICtrlCreateLabel( "", 33, 140, 394, 22 )
GUICtrlSetColor( -1, 0x0000ff )
Global $btnDatabase = GUICtrlCreateButton( "...", 431, 136, 24, 28 )
Global $lblEncyclopedia = GuiCtrlCreateLabel( "Active DMB encyclopedia: ", 30, 171, 365, 20 )
GUICtrlSetFont( -1, 10, 700, 0, $font )
Global $txtEncFrame = GUICtrlCreateLabel( "", 30, 196, 400, 28, $SS_ETCHEDFRAME )
Global $txtEncyclopedia = GUICtrlCreateLabel( "", 33, 200, 394, 22 )
GUICtrlSetColor( -1, 0x0000ff )
Global $btnEncyclopedia = GUICtrlCreateButton( "...", 431, 196, 24, 28 )

Global $btnStart = GUICtrlCreateButton( "Start", 288, 245, 85, 28 )
GUICtrlSetResizing( -1, $GUI_DOCKRIGHT + $GUI_DOCKSIZE )
GUICtrlSetState( -1, $GUI_DISABLE )
Global $btnExit = GUICtrlCreateButton( "Exit", 383, 245, 85, 28 )
GUICtrlSetResizing( -1, $GUI_DOCKRIGHT + $GUI_DOCKSIZE )
GUICtrlSetTip( -1, "Use Ctrl+Alt+X to quit while simming." )

; Add "Exit" item to program icon in notification area
Global $uxTrayExit = TrayCreateItem( "Exit" )
TraySetToolTip ( $SCRIPT_TITLE )

; Create status bar and set properties.
Global $frmStatusBar = _GUICtrlStatusBar_Create( $frmMain )
_GUICtrlStatusBar_SetSimple ( $frmStatusBar, True )
GUIRegisterMsg( $WM_SIZE, "WM_SIZE" )

; Set GUI Accelerator keys
Local $accelKeys[2][2] = [["{F1}", $mnuHelpAboutItem], ["{F5}", $mnuOptionsRefreshItem ]]
GUISetAccelerators( $accelKeys, $frmMain )

; Set the events.
GUICtrlSetOnEvent( $mnuFileDMBPathItem, "SetDMBPath" )
GUICtrlSetOnEvent( $mnuFileENCPathItem, "SetENCPath" )
GUICtrlSetOnEvent( $mnuFileExitItem, "ExitClicked" )
GUICtrlSetOnEvent( $mnuOptionsGBGStatsItem, "SetOptions" )
GUICtrlSetOnEvent( $mnuOptionsBoxscoreItem, "SetOptions" )
GUICtrlSetOnEvent( $mnuOptionsRefreshItem, "RefreshActiveDatabase" )
GUICtrlSetOnEvent( $mnuHelpAboutItem, "DisplayAbout" )
GUICtrlSetOnEvent( $txtSeasons, "SetSave" )
GUICtrlSetOnEvent( $txtStart, "SetSave" )
GUICtrlSetOnEvent( $btnDatabase, "SetDMBPath" )
GUICtrlSetOnEvent( $btnEncyclopedia, "SetENCPath" )
GUICtrlSetOnEvent( $btnStart, "StartButtonClicked" )
GUICtrlSetOnEvent( $btnExit, "ExitClicked" )
GUISetOnEvent( $GUI_EVENT_CLOSE, "ExitClicked" )
TrayItemSetOnEvent( $uxTrayExit, "ExitClicked" )

GUISetState()
#endregion --- GUI related code end ---

; Load startup settings from configuration file, if one exists.
; If Autosim.ini does not exist or can not be read, check for command line parameters.
; If there are no command line parameters, the defualt paths will be used.
Local $settings = IniReadSection( "autosim.ini", "StartupSettings" )
If @error = 1 Then
    GetCommandLineParams()
    UpdateSettings()
Else
    $g_dmbInstallPath = $settings[$DMB_PATH][1]
    $g_dmbEncInstallPath = $settings[$ENC_PATH][1]

    If $settings[$GBGSTATS][1] = "1" Then
        GUICtrlSetState( $mnuOptionsGBGStatsItem, $GUI_CHECKED )
        $g_startupSettings[$GBGSTATS][1] = "1"
    Else
        GUICtrlSetState( $mnuOptionsGBGStatsItem, $GUI_UNCHECKED )
        $g_startupSettings[$GBGSTATS][1] = "0"
    EndIf
    If $settings[$BOXSCORES][1] = "1" Then
        GUICtrlSetState( $mnuOptionsBoxscoreItem, $GUI_CHECKED )
        $g_startupSettings[$BOXSCORES][1] = "1"
    Else
        GUICtrlSetState( $mnuOptionsBoxscoreItem, $GUI_UNCHECKED )
        $g_startupSettings[$BOXSCORES][1] = "0"
    EndIf

    If $settings[$IS_RESTART][1] = "1" Then
        $g_startupSettings[$IS_RESTART][1] = "1"
        GUICtrlSetData( $txtSeasons, $settings[$LEFT][1] )
        GUICtrlSetData( $txtStart, $settings[$RESTART][1] )
        ; Update NumOfSeasons and StartAt to use regular settings next time AutoSim starts.
        $g_startupSettings[$SEASONS][1] = $settings[$SEASONS][1]
        $g_startupSettings[$START][1] = $settings[$START][1]
        SetSave()
    Else
        GUICtrlSetData( $txtSeasons, $settings[$SEASONS][1] )
        GUICtrlSetData( $txtStart, $settings[$START][1] )
        UpdateSettings()
    EndIf
EndIf

WinActivate( $frmMain )

;~ MsgBox(64, $SCRIPT_TITLE & " Configuration Settings", "DMBPath=" & $g_startupSettings[$DMB_PATH][1] & @LF & _
;~  "ENCPath=" & $g_startupSettings[$ENC_PATH][1] & @LF & _
;~  "NumOfSeasons=" & $g_startupSettings[$SEASONS][1] & @LF & _
;~  "StartAt=" & $g_startupSettings[$START][1] & @LF & _
;~  "IsRestart=" & $g_startupSettings[$IS_RESTART][1] & @LF & _
;~  "LeftToSim=" & $g_startupSettings[$LEFT][1] & @LF & _
;~  "RestartAt=" & $g_startupSettings[$RESTART][1])

While True
    Initialise()
    If @error Then
        Local $result = MsgBox( 17, $SCRIPT_TITLE, _
                                "Unable to locate Diamond Mind program files." & @CRLF & _
                                "Please update the paths to your DMB and DMBENC folders." )
        If $result = 1 Then
            SetDMBPath()
            SetENCPath()
        Else
            Exit(1)
        EndIf
    Else
        GUICtrlSetState( $btnStart, $GUI_ENABLE )
        ExitLoop
    EndIf
WEnd

While True
    Sleep( 1000 ) ; Wait for events.
WEnd

Exit

;------------------------------------------------------------------------------
; functions below

Func SetDMBPath()
    ;
    ; Return value:

    ; Prompt for folder location
    Local $dmbPath = FileSelectFolder( "Select the location of the DMB installation folder.", "" )

    ; Look for "baseball" application file to verify folder is a valid installation folder
    If @error Then
        MsgBox( 64, $SCRIPT_TITLE, "No folder was selected." )
    ElseIf FileExists( $dmbPath & "/baseball.exe" ) Then
        MsgBox( 64, $SCRIPT_TITLE, "The DMB installation folder is now set to: " & $dmbPath )
        $g_dmbInstallPath = $dmbPath
        SetSave()
        Initialise()
    Else
        MsgBox( 48, "DMB Installation Folder", "The folder '" & $dmbPath & _
                "' does not contain a DMB application file." )
    EndIf
    Return
EndFunc

Func SetENCPath()
    ;
    ; Return value:

    ; Prompt for folder location
    Local $encPath = FileSelectFolder( "Select the location of the DMB Encyclopedia installation folder.", "" )

    ; Look for "enc" application file to verify folder is a valid installation folder
    If @error Then
        MsgBox( 64, $SCRIPT_TITLE, "No folder was selected." )
    ElseIf FileExists( $encPath & "/enc.exe" ) Then
        MsgBox( 64, $SCRIPT_TITLE, "The DMBENC installation folder is now set to: " & $encPath )
        $g_dmbEncInstallPath = $encPath
        SetSave()
        Initialise()
    Else
        MsgBox( 48, "DMBENC Installation Folder", "The folder '" & $encPath & _
                "' does not contain a DMBENC application file." )
    EndIf
    Return
EndFunc

Func GetCommandLineParams()
    ; Check for program paths entered at command line.
    ; Return value: None
    If $CmdLine[0] > 0 Then
        $g_dmbInstallPath = $CmdLine[1]
        $g_dmbEncInstallPath = $CmdLine[2]
    EndIf
    Return 0
EndFunc

Func StartButtonClicked()
    ; Begin simulating seasons.
    ; Return value: None
    Local $dmbHandle, $dmbEncHandle
    Local $simTime = 0
    Local $importTime = 0
    Local $elapHr, $elapMin, $elapSec
    Local $remHr, $remMin, $remSec
    If IsValidSeason() = 1 And IsValidStart() = 1 Then
        Local const $NUM_SEASONS = GUICtrlRead( $txtSeasons )
        Local const $START_AT = GUICtrlRead( $txtStart )
    Else
        Return
    EndIf

    ; Disable program controls while DMB is running
    GUICtrlSetState( $btnStart, $GUI_DISABLE )
    GUICtrlSetState( $btnExit, $GUI_DISABLE )
    GUICtrlSetState( $mnuFileMenu, $GUI_DISABLE )
    GUICtrlSetState( $mnuOptionsMenu, $GUI_DISABLE )
    GUICtrlSetState( $mnuHelpMenu, $GUI_DISABLE )

    ; Run DMB and get it's Windows handle.
    Run( $g_dmbInstallPath & "\baseball.exe", $g_dmbInstallPath )
    WinWait( $DMB_TITLE )
    $dmbHandle = WinGetHandle( $DMB_TITLE )
    WinSetState( $DMB_TITLE, "", @SW_RESTORE )

    ; Run DMBEnc and get it's Windows handle.
    Run( $g_dmbEncInstallPath & "\enc.exe", $g_dmbEncInstallPath )
    WinWait( $DMBENC_TITLE  )
    $dmbEncHandle = WinGetHandle( $DMBENC_TITLE )
    WinSetState( $DMBENC_TITLE, "", @SW_RESTORE )

    ; Seed restart value. If DMB crashes restart info will be saved.
    $g_startupSettings[$IS_RESTART][1] = "1"

    ; Main loop to simulate seasons.
    For $i = 0 to $NUM_SEASONS - 1
        ; Status bar display e.g. Simming: 1005 (005/100) | Time: 75 min | Remaining: 105 min
        Local $statusMsg = StringFormat( "Simming: %04i (%03i/%03i)", _
                                         ($i + $START_AT), ($i + 1), $NUM_SEASONS )

        If $i > 0 Then
            _TicksToTime($simTime + $importTime, $elapHr, $elapMin, $elapSec)
            _TicksToTime((($simTime + $importTime )/$i) * ($NUM_SEASONS - $i), _
                         $remHr, $remMin, $remSec)
            $statusMsg &= StringFormat(" | Elapsed: %02ih:%02im | " & _
                                       "Remaining: %02ih:%02im", _
                                       $elapHr, $elapMin, $remHr, $remMin)
        EndIf

        ; Update status bar text
        _GUICtrlStatusBar_SetText( $frmStatusBar, $statusMsg, $SB_SIMPLEID )
        TraySetToolTip ( $statusMsg )
        $g_startupSettings[$LEFT][1] = ( $NUM_SEASONS - $i )
        $g_startupSettings[$RESTART][1] = ( $i + $START_AT )

        RestartSeason( $dmbHandle )

        $simTime = $simTime + SimulateSeason( $dmbHandle )
        $importTime = $importTime + ImportSeason( ( $i + $START_AT ), $dmbEncHandle )
    Next

    _GUICtrlStatusBar_SetText( $frmStatusBar, "", $SB_SIMPLEID )
    TraySetToolTip ( $SCRIPT_TITLE )
    $g_startupSettings[$IS_RESTART][1] = "0"

    RecapMsg( $simTime, $importTime, $NUM_SEASONS )

    ; Enable program controls
    GUICtrlSetState( $btnExit, $GUI_ENABLE )
    GUICtrlSetState( $mnuFileMenu, $GUI_ENABLE )
    GUICtrlSetState( $mnuOptionsMenu, $GUI_ENABLE )
    GUICtrlSetState( $mnuHelpMenu, $GUI_ENABLE )

    GUICtrlSetData( $txtSeasons, $g_startupSettings[$SEASONS][1] )
    GUICtrlSetData( $txtStart, $g_startupSettings[$START][1] )
    Initialise()
    If Not @error Then
        GUICtrlSetState( $btnStart, $GUI_ENABLE )
    EndIf
    Return
EndFunc

Func RefreshActiveDatabase()
    ; Refresh active db fields.
    ; Return value: None
    Initialise()
    If @error Then
        GUICtrlSetState( $btnStart, $GUI_DISABLE )
    Else
        GUICtrlSetState( $btnStart, $GUI_ENABLE )
        MsgBox( 64, $SCRIPT_TITLE, "The active database and encyclopedia files have been updated." )
    EndIf
    Return 0
EndFunc

Func SetOptions()
    ; Set persistent options in the Options menu
    If @GUI_CtrlId = $mnuOptionsGBGStatsItem Then
        If BitAND( GUICtrlRead($mnuOptionsGBGStatsItem), $GUI_CHECKED ) = $GUI_CHECKED Then
            GUICtrlSetState( $mnuOptionsGBGStatsItem, $GUI_UNCHECKED )
        Else
            GUICtrlSetState($mnuOptionsGBGStatsItem, $GUI_CHECKED)
        EndIf
    ElseIf @GUI_CtrlId = $mnuOptionsBoxscoreItem Then
        If BitAND( GUICtrlRead($mnuOptionsBoxscoreItem), $GUI_CHECKED ) = $GUI_CHECKED Then
            GUICtrlSetState( $mnuOptionsBoxscoreItem, $GUI_UNCHECKED )
        Else
            GUICtrlSetState($mnuOptionsBoxscoreItem, $GUI_CHECKED)
         EndIf
    EndIf
    SetSave()
EndFunc

Func IsValidSeason()
    ; Validate number of seasons.
    ; Return value: 1 - Valid number of seasons, 0 - Not valid.
    Local $NumOfSeasons = GUICtrlRead( $txtSeasons )
    If StringIsDigit( $NumOfSeasons ) = 0 Then
        MsgBox( 48, $SCRIPT_TITLE, "Number of seasons must be a positive number." )
        GUICtrlSetState( $txtSeasons, $GUI_FOCUS )
        Return 0
    EndIf
    Return 1
EndFunc

Func IsValidStart()
    ; Validate starting season.
    ; Return value: 1 - Valid starting season, 0 - Not valid.
    Local $startAt = GUICtrlRead( $txtStart )
    If StringIsDigit( $startAt ) = 0 Then
        MsgBox( 48, $SCRIPT_TITLE, "Starting season must be a positive number." )
        GUICtrlSetState( $txtStart, $GUI_FOCUS )
        Return 0
    EndIf
    Return 1
EndFunc

Func CloseAllRunningDMB()
    ; Close all open DMB or DMBEnc windows.
    ; Return value: None
    While WinExists( $DMB_TITLE )
        WinClose( $DMB_TITLE )
    WEnd
    Return 0
EndFunc

Func Initialise()
    ; Reads INI files to get state information at startup.
    ; Writes dmb source path to enc.ini.
    ; Return value: None
    Local $dmbDbNum = ""
    Local $dmbDbPath = ""
    Local $encDbNum = ""
    Local $encDbPath = ""
    CloseAllRunningDMB()
    ; Read baseball.ini to get numeric reference to active DMB Db.
    $dmbDbNum = IniRead( $g_dmbInstallPath & "\baseball.ini", "Locations", _
                         "ActivePlayerPath", "ERROR" )
    If $dmbDbNum <> "ERROR" Then
        ; Read baseball.ini to get path of active DMB Db.
        $dmbDbPath = IniRead( $g_dmbInstallPath & "\baseball.ini","Locations", _
                              "PlayerPath" & $dmbDbNum, "ERROR" )
        ; Read baseball.ini to get numeric reference to schedule.
        $g_schedule = IniRead( $g_dmbInstallPath & "\baseball.ini", "Locations", _
                               "LastLeague" & $dmbDbNum, "ERROR" )
        ; Read enc.ini to get numeric reference to active DMBEnc Db.
        $encDbNum = IniRead( $g_dmbEncInstallPath & "\enc.ini", "Locations", _
                             "ActiveEncPath", "ERROR" )
        If $encDbNum <> "ERROR" Then
            ; Read enc.ini to get path of active DMBEnc Db.
            $encDbPath = IniRead( $g_dmbEncInstallPath & "\enc.ini", "Locations", _
                                  "EncPath" & $encDbNum, "ERROR" )
            ; Update DMB source path in enc.ini to match active DMB Db path.
            IniWrite( $g_dmbEncInstallPath & "\enc.ini", "Locations", _
                      "DMBSrcPath", $dmbDbPath )
        Else
            SetError( 2 ) ; Error reading ActiveEncPath in enc.ini.
        EndIf
    Else
        SetError( 1 ) ; Error reading ActivePlayerPath in baseball.ini.
    EndIf
    If @error Then
        $g_schedule = 0 ; No schedule defined
        SetError( 1 )
    Else
        GUICtrlSetData( $txtDatabase, $dmbDbPath )
        GUICtrlSetData( $txtEncyclopedia, $encDbPath )
    EndIf
    Return 0
EndFunc

Func RestartSeason( $dmbHandle )
    ; Reset DMB season.
    ; Return value: None
    WinActivate( $dmbHandle )
    WinWaitActive( $dmbHandle )

    WinMenuSelectItem( $dmbHandle, "", "&Tools", "&Restart a season" )
    WinWait( "Season Restart Options", "Organization or league" )
    If $g_schedule > 0 Then
        ControlCommand( "Season Restart Options", "Organization or league", _
                        "ComboBox1", "SetCurrentSelection", $g_schedule - 1 )
    EndIf
    ControlClick( "Season Restart Options", "Organization or league", "Button1" )
    WinWaitClose( "Season Restart Options", "Organization or league" )

    WinWait( "Baseball", "Are you sure you want to reset" )
    ControlClick( "Baseball", "Are you sure you want to reset", "Button1" )
    WinWaitClose( "Baseball", "Are you sure you want to reset" )

    WinWait( "Baseball" )
    While WinExists( "Baseball", "Do you still want to reset the roster" )
        ControlClick( "Baseball", "Do you still want to reset the roster", "Button1" )
        WinWait( "Baseball" )
    WEnd

    WinWait( "Baseball", "Season has been restarted for" )
    ControlClick( "Baseball", "Season has been restarted for", "Button1" )
    WinWaitClose( "Baseball", "Season has been restarted for" )

    Return 0
EndFunc

Func SimulateSeason( $dmbHandle )
    ; Run season simulation.
    ; Return value: Time elapsed during simulation.

    Local $begin = TimerInit()
    WinActivate( $dmbHandle )
    WinWaitActive( $dmbHandle )

    WinMenuSelectItem( $dmbHandle, "", "&Game", "&Scheduled..." )
    WinWait( $DMB_TITLE, "Scheduled Game selection" )
    If $g_schedule > 0 Then
        ControlCommand( $DMB_TITLE, "Scheduled Game selection", "ComboBox1", _
                        "SetCurrentSelection", $g_schedule - 1 )
    EndIf

    ; Begin simming season
    WinMenuSelectItem( $dmbHandle, "Scheduled Game selection", "&Autoplay", "&All remaining games" )

    While True
        Sleep( 1000 )
        If WinExists( "Baseball", "The selected games have been completed" ) Then
            ; Capture end of simmed season.
            ControlClick( "Baseball", "The selected games have been completed", "Button1" )
            WinWaitClose( "Baseball", "The selected games have been completed" )

            WinWait( $DMB_TITLE, "Scheduled Game selection" )
            WinMenuSelectItem( $dmbHandle, "Scheduled Game selection", "&Game", "E&xit" )
            WinWaitClose( $DMB_TITLE, "Scheduled Game selection" )
            ExitLoop
        ElseIf Not WinExists( $dmbHandle ) Then
            MsgBox( 64, $SCRIPT_TITLE, "Diamond Mind has closed unexpectedly." & @CRLF & _
                    $SCRIPT_TITLE & " cannot continue. Your progress will be saved." )

            ; Database may be corrupted, not safe to continue
            ExitClicked()
        ElseIf WinExists( "BASEBALL Application" ) Then
        ; ControlClick( "BASEBALL Application", "&Close program", "Button1" )
            MsgBox( 64, $SCRIPT_TITLE, "Diamond Mind encountered an error." & @CRLF & _
                    $SCRIPT_TITLE & " cannot continue. Your progress will be saved." )

            ; Database may be corrupted, not safe to continue
            ExitClicked()
        ElseIf WinExists( "Baseball", "The computer manager was unable to field a valid" ) Then
            ; Close dialog
            ControlClick( "Baseball", "The computer manager was unable to field a valid", "Button1" )

            ; Capture games completed window
            WinWait( "Baseball", "The selected games have been completed" )
            ControlClick( "Baseball", "The selected games have been completed", "Button1" )
            WinWait( $DMB_TITLE, "Scheduled Game selection" )

            ; Restart simming
            WinMenuSelectItem( $DMB_TITLE, "Scheduled Game selection", "&Autoplay", "&All remaining games" )
        EndIf
    WEnd

    Return TimerDiff( $begin )    ; in milliseconds
EndFunc

Func ImportSeason( $year, $dmbEncHandle )
    ; Import season into Diamond Mind Baseball Encyclopedia.
    ; Return value: Time elapsed during import.
    Local $begin = TimerInit()
    WinActivate( $dmbEncHandle )
    WinWaitActive( $dmbEncHandle )

    WinMenuSelectItem( $dmbEncHandle, "", "&Tools", "&Import DMB season..." )
    WinWait( "Import season", "Location of Diamond Mind database" )
    ControlClick( "Import season", "Location of Diamond Mind database", "Button2" )
    WinWaitClose( "Import season", "Location of Diamond Mind database" )

    WinWait( "Season Add Options", "Organization or league" )
    If $g_schedule > 0 Then
        ControlCommand( "Season Add Options", "Organization or league", _
                        "ComboBox1", "SetCurrentSelection", $g_schedule - 1 )
    EndIf
    ControlSetText( "Season Add Options", "Year:", "Edit2", $year )

    If BitAND( GUICtrlRead($mnuOptionsGBGStatsItem), $GUI_CHECKED ) = $GUI_CHECKED Then
        ControlCommand( "Season Add Options", "", "[ID:3020]", "Check" )
    Else
        ControlCommand( "Season Add Options", "", "[ID:3020]", "UnCheck" )
    EndIf
    If BitAND( GUICtrlRead($mnuOptionsBoxscoreItem), $GUI_CHECKED ) = $GUI_CHECKED Then
        ControlCommand( "Season Add Options", "", "[ID:3021]", "Check" )
    Else
        ControlCommand( "Season Add Options", "", "[ID:3021]", "UnCheck" )
    EndIf

    ControlClick( "Season Add Options", "Organization or league", "Button3" )
    WinWaitClose( "Season Add Options", "Organization or league" )

    While True
        Sleep( 1000 )
        If WinExists( "Enc", "Season successfully loaded into the encyclopedia" ) Then
            ControlClick( "Enc", "Season successfully loaded into the encyclopedia", "Button1" )
            WinWaitClose( "Enc", "Season successfully loaded into the encyclopedia" )
            ExitLoop
        ElseIf WinExists( "Enc", "Unable to load season" ) Then
            MsgBox( 64, $SCRIPT_TITLE, "Encyclopedia was unable to import the season." & @CRLF & _
                    $SCRIPT_TITLE & " cannot continue. Your progress will be saved." )
            ExitClicked()
		 EndIf
	  WEnd

    Return TimerDiff( $begin )    ; in milliseconds
EndFunc

Func RecapMsg( $simTime, $importTime, $numOfSeasons )
    ; Display dialog with batch run statistics
    ; Return value: None
    Local $simHr, $simMin, $simSec
    Local $simAveHr, $simAveMin, $simAveSec
    Local $impHr, $impMin, $impSec
    Local $impAveHr, $impAveMin, $impAveSec

    _TicksToTime( $simTime, $simHr, $simMin, $simSec )
    _TicksToTime( $simTime/$numOfSeasons, $simAveHr, $simAveMin, $simAveSec )
    _TicksToTime( $importTime, $impHr, $impMin, $impSec )
    _TicksToTime( $importTime/$numOfSeasons, _
                  $impAveHr, $impAveMin, $impAveSec )

    Local $reportText = "Seasons simmed: " & $numOfSeasons & @CRLF & @CRLF & _
        StringFormat( "Sim time: %02ih:%02im:%02is", _
                      $simHr, $simMin, $simSec ) & @CRLF & _
        StringFormat( "Sim average: %02ih:%02im:%02is", _
                      $simAveHr, $simAveMin, $simAveSec ) & @CRLF & @CRLF & _
        StringFormat( "Import time: %02ih:%02im:%02is", _
                      $impHr, $impMin, $impSec ) & @CRLF & _
        StringFormat( "Import average: %02ih:%02im:%02is", _
                      $impAveHr, $impAveMin, $impAveSec )
    MsgBox( 64, $SCRIPT_TITLE & " Summary", $reportText )
    Return 0
EndFunc

Func WM_SIZE( $hWnd )
    ; Resize the status bar when GUI size changes
    ; Return value:
    _GUICtrlStatusBar_Resize ( $frmStatusBar )
    Return $GUI_RUNDEFMSG
EndFunc

Func SetSave()
    ; Set flag to save settings to config file on exit.
    ; Return value: None
    $g_saveOnExit = 1
    UpdateSettings()
    Return
EndFunc

Func UpdateSettings()
    ; Update startup settings array.
    ; Return value: None
    $g_startupSettings[$DMB_PATH][1] = $g_dmbInstallPath
    $g_startupSettings[$ENC_PATH][1] = $g_dmbEncInstallPath

    If BitAND( GUICtrlRead($mnuOptionsGBGStatsItem), $GUI_CHECKED ) = $GUI_CHECKED Then
        $g_startupSettings[$GBGSTATS][1] = "1"
    Else
        $g_startupSettings[$GBGSTATS][1] = "0"
    EndIf
    If BitAND( GUICtrlRead($mnuOptionsBoxscoreItem), $GUI_CHECKED ) = $GUI_CHECKED Then
        $g_startupSettings[$BOXSCORES][1] = "1"
    Else
        $g_startupSettings[$BOXSCORES][1] = "0"
    EndIf

    If $g_startupSettings[$IS_RESTART][1] <> "1" Then
        $g_startupSettings[$SEASONS][1] = GUICtrlRead( $txtSeasons )
        $g_startupSettings[$START][1] = GUICtrlRead( $txtStart )
    EndIf
    $g_startupSettings[$IS_RESTART][1] = "0"
    $g_startupSettings[$LEFT][1] = "0"
    $g_startupSettings[$RESTART][1] = "0"
    Return
EndFunc

Func DisplayAbout()
    ; Displays an About message.
    ; Return value: None
    Local $reportText = $SCRIPT_TITLE & " " & $SCRIPT_VERSION & @CRLF & _
                        "December 2021" & @CRLF & @CRLF & _
                        "Copyright (C) 2022 David Pyke." & @CRLF & _
                        "https://github.com/fishinnabarrel/AutoSim"
    MsgBox( 0, "About " & $SCRIPT_TITLE, $reportText )
    Return 0
EndFunc

Func displayHelp()
    ; Displays a help message.
    ; Return value: None
    Return 0
EndFunc

Func ExitClicked()
    ; Exits script in response to Exit Event.
    ; Return vale: None

    ; If settings have been updated, save startup settings to configuration file - autosim.ini.
    If $g_saveOnExit Or $g_startupSettings[$IS_RESTART][1] = "1" Then
        Local $settingsText = "DMBPath=" & $g_startupSettings[$DMB_PATH][1] & @LF & _
            "ENCPath=" & $g_startupSettings[$ENC_PATH][1] & @LF & _
            "NumOfSeasons=" & $g_startupSettings[$SEASONS][1] & @LF & _
            "StartAt=" & $g_startupSettings[$START][1] & @LF & _
            "IsRestart=" & $g_startupSettings[$IS_RESTART][1] & @LF & _
            "LeftToSim=" & $g_startupSettings[$LEFT][1] & @LF & _
            "RestartAt=" & $g_startupSettings[$RESTART][1] & @LF & _
            "SaveGBGStats=" & $g_startupSettings[$GBGSTATS][1] & @LF & _
            "SaveBoxscores=" & $g_startupSettings[$BOXSCORES][1]
        IniWriteSection( "autosim.ini", "StartupSettings", $settingsText )
    EndIf
    Exit
EndFunc
