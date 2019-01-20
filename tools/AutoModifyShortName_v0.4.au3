; ------------------------------------------------------------------------
;
; Title:          AutoModifyShortName.au3
; Developers:     David Pyke
; Date:           June 10, 2007
; Modified:       July 12, 2016
; Version:        0.4
;
; Script Function: Automate player modification - change short name.
; ------------------------------------------------------------------------

;#RequireAdmin
#include <Debug.au3>
#include <Misc.au3>

Opt("MustDeclareVars", 1)   ; 0=no, 1=require pre-declare
Opt("WinTextMatchMode", 2)  ; 1=complete, 2=quick
Opt("WinTitleMatchMode", 4) ; 1=start, 2=subStr, 3=exact, 4=advanced

Global Const $ScriptTitle = "AutoModifyShortName v0.4"
Global Const $DMBTitle = "Diamond Mind Baseball"

; Set length of name fields used to build short name
; total must be less than 15
Local $lastNameLen = 13
Local $firstNameLen = 1

; Check if script is already running.
_Singleton(@ScriptName, 0)

HotKeySet("^!x", "exitEarly")  ; Ctrl-Alt-x to exit script

TraySetToolTip($ScriptTitle)

; Check if DMB is running.
If Not WinExists($DMBTitle) Then
    MsgBox(48, $ScriptTitle, "Diamond Mind Baseball is not open!" & @LF & _
                                                        "Script will now exit.")
    Exit(1)
EndIf

OpenPlayer()
Local $lastPID = -1

; Change the short name for the first player selected.
Local $currentPID = ModifyPlayer($lastNameLen, $firstNameLen)

; Stop when the last player is reached.
While $currentPID <> $lastPID
    $lastPID = $currentPID
    Send("{DOWN}")
    $currentPID = ModifyPlayer($lastNameLen, $firstNameLen)
WEnd

MsgBox(0, $ScriptTitle, "Task complete.")

Exit(0)

; ----------------------------------------------------------------------------
; functions below

Func OpenPlayer()
    ; ; Open organizer window and select Player tab
    ; Returns: None

    WinActivate($DMBTitle, "")
    WinWaitActive($DMBTitle, "")
    WinSetState($DMBTitle, "", @SW_MAXIMIZE)
    WinMenuSelectItem($DMBTitle, "","&View", "&Organizer...")
    WinWait($DMBTitle, "Organizer")
;~     WinMenuSelectItem($DMBTitle, "","&Window", "&Tile")
    ControlClick($DMBTitle, "Organizer", "[ID:59905]", "left", 1, 175, 8)

    Return
EndFunc

Func ModifyPlayer(Const $lastNameLen, Const $firstNameLen)
    ; Modifies player short name in the form - "Last,F"
    ; For example, Aaron,H
    ; Returns: UID of player modified

    _Assert(($lastNameLen + $firstNameLen) < 15, True, 1, @ScriptLineNumber)

    ; Open Modify window and select General player info
    WinWait($DMBTitle, "Organizer")
    ControlClick($DMBTitle, "Organizer", "Modify")
    Send("g")
    WinWait("Modify Player")

    ; Modify short name if no custom short name (i.e. no comma in field)
    Local $pid = ControlGetText("Modify Player", "", "[ID:1380]")
    Local $lastName = ControlGetText("Modify Player", "", "[ID:1373]")
    Local $shortName = ControlGetText("Modify Player", "", "[ID:1783]")
    If $lastName == $shortName Then
        Local $firstName = ControlGetText("Modify Player", "", "[ID:1372]")
        $shortName = StringLeft($lastName, $lastNameLen) & _
                        "," & StringLeft($firstName, $firstNameLen)
        ControlSetText("Modify Player", "", "[ID:1783]", $shortName)
        Sleep(100)
    ElseIf Not StringInStr($shortName, ",") Then
        Local $firstName = ControlGetText("Modify Player", "", "[ID:1372]")
        $shortName &= "," & StringLeft($firstName, 1)
        ControlSetText("Modify Player", "", "[ID:1783]", $shortName)
        Sleep(100)
    EndIf

    ControlClick("Modify Player", "", "OK")
    WinWait($DMBTitle, "Organizer")

    Return $pid
EndFunc

Func exitEarly()
    ; Exists script early - from hotkey.

    Exit(1)
EndFunc
