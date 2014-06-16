#SingleInstance

;**************************************************************************
; NOTE TO READERS:
;
; Like my vimrc file, this is the result of years of just quickly dumping
; stuff into this file just to get things done... some of the scripts were 
; one-off solutions to long forgotten problems, some were interesting ideas
; that never worked, etc... 
;
; Autohotkey has an "interesting" syntax for manipulating numbers vs. strings
; which always seemed to trip me up, so some of the math functions might
; be more complicated than they need to be -- don't learn from them. :-)
;





;**************************************************************************
; Function Documentation                                                {{{
;
; <CAS>A -- 
; < AS>A -- 
; <C S>A --
; <CA >A --
;
; <CAS>B -- Search for clip<B>oard content
; < AS>B --
; <C S>B --
; <CA >B --
;
; <CAS>C -- Cycle Command Prompts (cygwin/4NT/cmd)
; < AS>C -- Open Cygwin Prompt
; <C S>C -- Open Command Prompt (4NT)
; <CA >C -- Launch/Activate Outlook Calendar
;
; <CAS>D -- Cycle DevStudio Windows
; <C S>D -- Launch DevStudio
; <CA >D -- Send <winkey><d> (toggle show desktop)
; < AS>D -- 
;
; <CAS>E -- Cycle EA Windows
; <C S>E -- Launch EA
; <CA >E -- Send <winkey><e> (launch explorer)
; < AS>E --
;
; <CAS>F -- Cycle Firefox Windows
; <C S>F --
; <CA >F -- Launch SSH forwarding connection
; < AS>F --
;
; <CAS>G -- Cycle Google Chrome Windows
; <C S>G -- Launch Google Chrome
; <CA >G --
; < AS>G --
;
; <CAS>H -- Switch between Putty Windows
; <C S>H -- Log into home computer using Putty
; <CA >H --
; < AS>H --
;
; <CAS>I -- Cycle between Safari Windows
; <C S>I --
; <CA >I --
; < AS>I --
;
; <CAS>J -- Volume Down
; <C S>J -- Eject USB devices
; <CA >J -- Eject DVD drive
; < AS>J --
;
; <CAS>K -- Volume Up
; <C S>K -- 
; <CA >K --
; < AS>K --
;
; <CAS>L -- Cycle between explorer windows
; <C S>L -- 
; <CA >L -- Send <winkey><l> (lock terminal)
; < AS>L -- Listen for client script
;
; <CAS>M -- Mute
; <C S>M -- New Outlook Message
; <CA >M -- Cycle through outlook messages
; < AS>M --
;
; <CAS>N --
; <C S>N -- 
; <CA >N --
; < AS>N --
;
; <CAS>O -- Launch/Activate Outlook
; <C S>O -- 
; <CA >O --
; < AS>O --
;
; <CAS>P -- Media Play/Pause
; <C S>P -- 
; <CA >P --
; < AS>P --
;
; <CAS>Q --
; <C S>Q -- 
; <CA >Q --
; < AS>Q --
;
; <CAS>R --
; <C S>R -- 
; <CA >R -- Send <winkey><r> (run program)
; < AS>R --
;
; <CAS>S -- Media Stop
; <C S>S -- Launch SCM tool GUI
; <CA >S -- Switch to SCM tool GUI window
; < AS>S --
;
; <CAS>T -- Outlook Tasks
; <C S>T -- 
; <CA >T --
; < AS>T --
;
; <CAS>U -- Activate Ubuntu VM
; <C S>U -- Launch Ubuntu VM
; <CA >U --
; < AS>U --
;
; <CAS>V -- Cycle between Vim windows
; <C S>V -- Launch Vim
; <CA >V --
; < AS>V --
;
; <CAS>W -- Cycle between MS Word windows
; <C S>W -- Launch MS Word, new document
; <CA >W --
; < AS>W --
;
; <CAS>X -- Cycle between Excel windows
; <C S>X -- Launch MS Excel, new spreadsheet
; <CA >X --
; < AS>X -- Launch X server
;
; <CAS>Y --
; <C S>Y -- 
; <CA >Y --
; < AS>Y --
;
; <CAS>Z --
; <C S>Z -- 
; <CA >Z --
; < AS>Z --
;
; <CAS>, -- Previous track
; <C S>, -- 
; <CA >, --
; < AS>, --
;
; <CAS>. -- Next track
; <C S>. -- 
; <CA >. --
; < AS>. --
;                                                                       }}}
;**************************************************************************
; TIMERS                                                                {{{
;**************************************************************************
SetTimer,KillAutoUpdatesDlg,500
SetTimer,UPDATEDSCRIPT,5000
SetTimer,UPDATEDCONFIG,5000
;                                                                       }}}
;**************************************************************************
; WINDOW CLASS FOR WINDOW CLASS VARIABLES THAT CHANGE                   {{{
;**************************************************************************
EA_WINCLASS = Afx:00400000:b:00010013:00000006:01A90B95
RHAPSODY_WINCLASS = TEST_WIN32WND
;                                                                       }}}
;**************************************************************************
; Read Configuration Values                                             {{{
;**************************************************************************
config_file = %A_ScriptDir%\ahk.personal.rc
b_haschrome:=false
b_hasfirefox:=false
b_hasea:=false
b_hasscm:=false
b_hasmsword:=false
b_hasexcel:=false
b_hasvisualstudio:=false
b_useVolumeOSD:=true
b_preventScreenSaver:=false
s_actual_script = %A_ScriptFullPath%

; Default to Windows Media Player if a default player isn't specfied
defaultmediaplayercommand=wmplayer

; Default to notepad (cringe!) if the vim command isn't specified
vimcommand=notepad

; Fallback browser if one of the prefered browsers isn't configured
fallbackbrowser=IEXPLORE.EXE

; Default monitor positions in case there isn't a config file
leftmonitor=1
rightmonitor=2

; Timeout value for double paste for plaintext
pasteagaintimeout=200

; Audio controls client/server configuration
b_isserver:=false
b_isclient:=false
ServerIP:="127.0.0.1"
CommPort:=8765
WinMessageID:= 0x5555

Loop, Read, %config_file%
{
   ; Ignore commented out lines (; at column 0)
   StringGetPos, firstchar, A_LoopReadLine, `;
   if firstchar = 0
   {
      continue
   }

   ; Ignore blank lines
   if A_LoopReadLine =  
   {
      continue
   }
   
   IfInString, A_LoopReadLine, actualscript
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     s_actual_script = %f_line2%
   }
   else IfInString, A_LoopReadLine, leftmonitor
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     leftmonitor = %f_line2%
   }
   else IfInString, A_LoopReadLine, rightmonitor
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     rightmonitor = %f_line2%
   }
   else IfInString, A_LoopReadLine, vimcommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     vimcommand = %f_line2%
   }
   else IfInString, A_LoopReadLine, scmcommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     scmcommand = %f_line2%
     b_hasscm:=true
   }
   else IfInString, A_LoopReadLine, visualstudiocommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     visualstudiocommand = %f_line2%
     b_hasvisualstudio:=true
   }
   else IfInString, A_LoopReadLine, eacommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     eacommand = %f_line2%
     b_hasea:=true
   }
   else IfInString, A_LoopReadLine, chromecommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     chromecommand = %f_line2%
     b_haschrome:=true
   }
   else IfInString, A_LoopReadLine, firefoxcommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     firefoxcommand = %f_line2%
     b_hasfirefox:=true
   }
   else IfInString, A_LoopReadLine, mswordcommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     mswordcommand = %f_line2%
     b_hasmsword:=true
   }
   else IfInString, A_LoopReadLine, excelcommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     excelcommand = %f_line2%
     b_hasexcel:=true
   }
   else IfInString, A_LoopReadLine, defaultmediaplayercommand
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     defaultmediaplayercommand = %f_line2%
   }
   else IfInString, A_LoopReadLine, pasteagaintimeout
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     pasteagaintimeout = %f_line2%
   }
   else IfInString, A_LoopReadLine, serverip
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     ServerIP = %f_line2%
   }
   else IfInString, A_LoopReadLine, commport
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     CommPort = %f_line2%
   }
   else IfInString, A_LoopReadLine, winmessageid
   {
     StringSplit, f_line, A_LoopReadLine, `=
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     WinMessageID = %f_line2%
   }
   else IfInString, A_LoopReadLine, isclient
   {
     b_isclient:=true
   }
   else IfInString, A_LoopReadLine, isserver
   {
     b_isserver:=true
   }

   ; Misc Configuration Stuff
   else IfInString, A_LoopReadLine, preventScreenSaver
   {
      b_preventScreenSaver:=true
   }
   else IfInString, A_LoopReadLine, preventVolumeOSD
   {
     b_useVolumeOSD:=false
   }
   else
   {
      MsgBox, Error parsing %config_file%, unable to interpret line:`n`n%A_LoopReadLine%
   }
}

IfNotExist, %s_actual_script%
{
   MsgBox, 16, ERROR, 'actualscript' setting points to an invalid file:  %s_actual_script%
}

FileGetSize, config_file_exists, %config_file%
if config_file_exists =
{
   MsgBox, Creating default %config_file%
   FileAppend, `;******************************************************************`n,%config_file%
   FileAppend, `;          Section Used For Configuration of EXE locations         `n,%config_file%
   FileAppend, `;******************************************************************`n,%config_file%
   FileAppend, `;defaultmediaplayercommand=`n,%config_file%
   FileAppend, `;eacommand=`n,%config_file%
   FileAppend, `;excelcommand=`n,%config_file%
   FileAppend, `;chromecommand=`n,%config_file%
   FileAppend, `;firefoxcommand=`n,%config_file%
   FileAppend, `;mswordcommand=`n,%config_file%
   FileAppend, `;scmcommand=`n,%config_file%
   FileAppend, `;vimcommand=`n,%config_file%
   FileAppend, `;visualstudiocommand=`n,%config_file%
   FileAppend, `n,%config_file%
   FileAppend, `;******************************************************************`n,%config_file%
   FileAppend, `;     Section Used For Configuration of Client/Server hotkeys      `n,%config_file%
   FileAppend, `;******************************************************************`n,%config_file%
   FileAppend, `;serverip=volrcd168`n,%config_file%
   FileAppend, `;commport=8765`n,%config_file%
   FileAppend, `;winmessageid=0x5555`n,%config_file%
   FileAppend, `;isserver`n,%config_file%
   FileAppend, `;isclient`n,%config_file%
   FileAppend, `n,%config_file%
   FileAppend, `n,%config_file%
   FileAppend, `;******************************************************************`n,%config_file%
   FileAppend, `;                 Miscelaneous Configuration Stuff                 `n,%config_file%
   FileAppend, `;******************************************************************`n,%config_file%
   FileAppend, `;pasteagaintimeout=200`n,%config_file%
   FileAppend, `;preventScreenSaver`n,%config_file%
   FileAppend, `;preventVolumeOSD`n,%config_file%
   FileAppend, `;leftmonitor=1`n, %config_file%
   FileAppend, `;rightmonitor=2`n, %config_file%
   FileAppend, `;actualscript=`n, %config_file%
   FileAppend, `n,%config_file%
   FileAppend, `; Last line comment since AHK doesn't seem to handle a non terminated`n, %config_file%
   FileAppend, `; last line... this'll prevent accidental problems`n, %config_file%
}
;                                                                }}}
;**************************************************************************



if(b_preventScreenSaver)
{
   SetTimer, PREVENTSCREENSAVER, 10000
}

if(b_isserver)
{ 
  TrayTip, Rex's Macros, Configured as AHK server for remote commands on port %commport%, 5, 1
  OnExit, ShutdownClientServer
}

if(b_isclient)
{
  OnExit, ShutdownClientServer

  socket := ConnectToAddress(serverip, commport)
  if socket = -1
  {
     MsgBox, Unable to connect to AHK server!
  }
}

;*****************************************************************************
; Autoexecute section for Paste Plain functionality                        {{{
;*****************************************************************************
pastecounter=0
Hotkey,$^v,PASTEONCE,On


;                                                                          }}}
;*****************************************************************************
; Autoexecute section for Favorites Folder functionality                   {{{
;*****************************************************************************
Hotkey, ^!+F1, f_DisplayMenu

/*
ITEMS IN FAVORITES MENU <-- Do not change this string.
Desktop       ; %A_Desktop%
&My Documents ; %A_MyDocuments%
&Program Files; %A_ProgramFiles%

&C:\          ; C:\
&D:\          ; D:\
*/

;----Read the configuration file.
f_AtStartingPos = n
f_MenuItemCount = 0
Loop, Read, %s_actual_script%
{
   ; Skip over all lines until the starting line is arrived at.
   if f_AtStartingPos = n
   {
     IfInString, A_LoopReadLine, ITEMS IN FAVORITES MENU
     {
        f_AtStartingPos = y
     }
     continue  ; Start a new loop iteration.
   }
   
   ; Otherwise, the closing comment symbol marks the end of the list.
   if A_LoopReadLine = */
     break  ; terminate the loop
  
   ; Menu separator lines must also be counted to be compatible
   ; with A_ThisMenuItemPos:
   f_MenuItemCount++
   if A_LoopReadLine =  ; Blank indicates a separator line.
     Menu, Favorites, Add
   else
   {
     StringSplit, f_line, A_LoopReadLine, `;
     f_line1 = %f_line1%  ; Trim leading and trailing spaces.
     f_line2 = %f_line2%  ; Trim leading and trailing spaces.
     ; Resolve any references to variables within either field, and
     ; create a new array element containing the path of this favorite:
     Transform, f_path%f_MenuItemCount%, deref, %f_line2%
     Transform, f_line1, deref, %f_line1%
     Menu, Favorites, Add, %f_line1%, f_OpenFavorite
   }
}



;                                                                          }}}
;*****************************************************************************
; Create Menus                                                          {{{
;
; Search for clipboard contents autoexecute
Menu, SearchTypes, Add, Open in &Internet Explorer, HandleClipboard
Menu, SearchTypes, Add, &Google, HandleClipboard
Menu, SearchTypes, Add, Google G&DS, HandleClipboard
Menu, SearchTypes, Add, &Wikipedia, HandleClipboard
Menu, SearchTypes, Add, &Merriam-Webster, HandleClipboard
Menu, SearchTypes, Add
Menu, SearchTypes, Add, EA -- &Non-Functional Requirement, HandleClipboard
Menu, SearchTypes, Add, EA -- &Functional Requirement, HandleClipboard
Menu, SearchTypes, Add
Menu, SearchTypes, Add, &Run Contents, RunClipboard
if(b_isclient)
{
   Menu, SearchTypes, Add, &Send Contents, SendClipboard
}


;  Multi Launcher
Menu, MLMenu, Add, &Firefox, MultiLauncherFunction
Menu, MLMenu, Add, &Calc, MultiLauncherFunction
Menu, MLMenu, Add
Menu, MLMenu, Add, &Run Clipboard Contents, RunClipboard

;*****************************************************************************
; Misc Commands
;*****************************************************************************
Menu, MiscCommandsMenu, Add, Close All E&xplorer Windows, MiscCommandsFunction
Menu, MiscCommandsMenu, Add, Close All E-Mail &Message Windows, MiscCommandsFunction
Menu, MiscCommandsMenu, Add, Update &Doxyments, MiscCommandsFunction

;*****************************************************************************
; Window Positions
;*****************************************************************************
Menu, LeftMonitorWindowPositionsMenu, Add, Whole &Monitor, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add
Menu, LeftMonitorWindowPositionsMenu, Add, &Left, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add, &Right, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add, &Top, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add, &Bottom, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add
Menu, LeftMonitorWindowPositionsMenu, Add, Quadrant &1, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add, Quadrant &2, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add, Quadrant &3, WindowPositionsFunction
Menu, LeftMonitorWindowPositionsMenu, Add, Quadrant &4, WindowPositionsFunction

Menu, RightMonitorWindowPositionsMenu, Add, Whole &Monitor, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add
Menu, RightMonitorWindowPositionsMenu, Add, &Left, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add, &Right, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add, &Top, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add, &Bottom, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add
Menu, RightMonitorWindowPositionsMenu, Add, Quadrant &1, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add, Quadrant &2, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add, Quadrant &3, WindowPositionsFunction
Menu, RightMonitorWindowPositionsMenu, Add, Quadrant &4, WindowPositionsFunction

Menu, ColumnWindowPositionsMenu, Add, &1, WindowPositionsFunction
Menu, ColumnWindowPositionsMenu, Add, &2, WindowPositionsFunction
Menu, ColumnWindowPositionsMenu, Add, &3, WindowPositionsFunction
Menu, ColumnWindowPositionsMenu, Add, &4, WindowPositionsFunction
Menu, ColumnWindowPositionsMenu, Add, &5, WindowPositionsFunction
Menu, ColumnWindowPositionsMenu, Add, &6, WindowPositionsFunction

Menu, WindowPositionsMenu, Add, &Left, :LeftMonitorWindowPositionsMenu
Menu, WindowPositionsMenu, Add, &Right, :RightMonitorWindowPositionsMenu
Menu, WindowPositionsMenu, Add, &Columns, :ColumnWindowPositionsMenu
Menu, WindowPositionsMenu, Add
Menu, WindowPositionsMenu, Add, &Switch, WindowPositionsFunction

Menu, LyncMeetingsMenu, Add, PRIVATE, LyncMeetingsFunction

;*****************************************************************************
;    END OF AUTORUN, INCLUDING MACHINE SPECIFIC FILE WHICH MIGHT END IT
;*****************************************************************************
#Include %A_ScriptDir%\this_machine.ahk

^!+b:: 
Menu, SearchTypes, show
return

^!+F2:: 
Menu, MLMenu, show
return

^!+F5::
Menu, MiscCommandsMenu, show
return

^!+F6::
Menu, LyncMeetingsMenu, show
return

^!+F12::
Menu, WindowPositionsMenu, show
return


;                                                                       }}}


;*****************************************************************************
; <CAS><DEL> Toggle Script ON/OFF {{{
^!+Delete::Suspend
; }}}
;*****************************************************************************

;*****************************************************************************
WindowPositionsFunction:                                  ;{{{
{   
   if( A_ThisMenu = "WindowPositionsMenu" && A_ThisMenuItemPos = 5)
   {
      GoSub, SwitchMonitors
      return
   }

   if( A_ThisMenu = "LeftMonitorWindowPositionsMenu")
   {
      monnum := leftmonitor
   }
   else if( A_ThisMenu = "RightMonitorWindowPositionsMenu")
   {
      monnum := rightmonitor
   }
   else if( A_ThisMenu = "ColumnWindowPositionsMenu")
   {
      UseColumn( A_ThisMenuItemPos )
      return
   }
   else
   {
      MsgBox, Ooops... AHK script error!
   }

   if(A_ThisMenuItemPos = 1)       ; Whole Monitor
   {
      UseMonitor(monnum)
   }
   else if(A_ThisMenuItemPos = 3)  ; Left Half
   {
      UseLeftHalfOfMonitor(monnum)
   }
   else if(A_ThisMenuItemPos = 4)  ; Right Half
   {
      UseRightHalfOfMonitor(monnum)
   }
   else if(A_ThisMenuItemPos = 5)  ; Top Half
   {
      UseTopHalfOfMonitor(monnum)
   }
   else if(A_ThisMenuItemPos = 6)  ; Bottom Half
   {
      UseBottomHalfOfMonitor(monnum)
   }
   else if(A_ThisMenuItemPos = 8)  
   {
      UseQuadrantOfMonitor(monnum, 1)
   }
   else if(A_ThisMenuItemPos = 9)
   {
      UseQuadrantOfMonitor(monnum, 2)
   }
   else if(A_ThisMenuItemPos = 10)
   {
      UseQuadrantOfMonitor(monnum, 3)
   }
   else if(A_ThisMenuItemPos = 11)
   {
      UseQuadrantOfMonitor(monnum, 4)
   }
   else
   {
      MsgBox, Ooops... AHK script error!
   }
}
return
;}}}
;*****************************************************************************



;*****************************************************************************
LyncMeetingsFunction: ;{{{
   ; PRIVATE
return
;}}}
;*****************************************************************************




;*****************************************************************************
;  Windows Key Replacement Mappings                                        {{{
;*****************************************************************************

;------------------------------------------------
;--------- Remap Replacement for Win-P ----------
;------------------------------------------------
^!p::
Send, #p
return

;------------------------------------------------
;--------- Remap Replacement for Win-D ----------
;------------------------------------------------
^!d::
Send, #d
return

;------------------------------------------------
;--------- Remap Replacement for Win-E ----------
;------------------------------------------------
^!e::
Send, #e
return

; Cycle between explorer windows
^!+l::
   GroupAdd, explorer_windows, ahk_class CabinetWClass
   GroupAdd, explorer_windows, ahk_class ExploreWClass
   GroupActivate, explorer_windows, R
return


;------------------------------------------------
;--------- Remap Replacement for Win-L ----------
;------------------------------------------------
^!l::
   DllCall("LockWorkStation")
return

;------------------------------------------------
;--------- Remap Replacement for Win-R ----------
;------------------------------------------------
^!r::
   Send, #r
   WinWait, ahk_class #32770 
   WinActivate
return
;------------------------------------------------


;------------------------------------------------
;------- Remap Replacement for Win-Break --------
;------------------------------------------------
^!SC146::
   Run, C:\WINDOWS\system32\sysdm.cpl
return
;------------------------------------------------

;------------------------------------------------
;--------------- Open Desktop CPL ---------------
;------------------------------------------------
^!+SC146::
Run, C:\WINDOWS\system32\desk.cpl
return
;------------------------------------------------


;                                                                          }}}
;*****************************************************************************


;*****************************************************************************
; <A>` -- Cycle windows of the same winclass as the current window 
;              (like Linux window managers does)
;*****************************************************************************
!`::
  WinGetClass, currentwinclass, A
  ; There is no way to delete or clear groups, so create a unique group per
  ; winclass.  This is a little bit wasteful, as it potentially leaves behind a
  ; lot of unnecessary groups, but given that it will probably be used with a
  ; limited set of windows per run of the script, I can't imagine it leaking
  ; too many resources...
  groupname:="cycle_current_win_"+currentwinclass  
  GroupAdd, %groupname%, ahk_class %currentwinclass%
  GroupActivate, %groupname%, R
return


;*****************************************************************************
; <CAS>N -- Cycle Lync Windows
;*****************************************************************************
^!+n::
   GroupAdd, lync_windows, ahk_class IMWindowClass 
   GroupAdd, lync_windows, ahk_class CommunicatorMainWindowClass
   GroupActivate, lync_windows, R
return

;*****************************************************************************
; <CS>N -- Open Lync main window
;*****************************************************************************
;^+n::
;   Run, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Lync\Microsoft Lync 2010.lnk"
;return


;*****************************************************************************
;**                                                                         **
;**                           Outlook Keymappings                           **
;**                                                                         **
;*****************************************************************************

CreateOutlookMessageGroup:
   SetTitleMatchMode, 2
   GroupAdd, msg_windows, Untitled Message
   GroupAdd, msg_windows, - Message
   GroupAdd, msg_windows, - Meeting
return

;------------------------------------------------
;          <CAS>O -- Launch/Activate Outlook Mail
;------------------------------------------------
^!+o::
   Run, C:\Program Files (x86)\Microsoft Office\OFFICE14\OUTLOOK.EXE  /recycle
return

;------------------------------------------------
;         <CAS>T -- Launch/Activate Outlook Tasks
;------------------------------------------------
^!+t::
   Run, C:\Program Files (x86)\Microsoft Office\OFFICE14\OUTLOOK.EXE /recycle /select outlook:tasks
return

;------------------------------------------------
;                <CS>T -- Create New Outlook Task
;------------------------------------------------
; ^+t::
;    Run, C:\Program Files (x86)\Microsoft Office\OFFICE14\OUTLOOK.EXE /c ipm.task
; return

;------------------------------------------------
;       <CA>C -- Launch/Activate Outlook Calendar
;------------------------------------------------
^!C:: ; Launch calendar
IfWinExist, Calendar - Mailbox - Kerr
{
  WinActivate
}
else
{
  Run, C:\Program Files (x86)\Microsoft Office\OFFICE14\OUTLOOK.EXE /recycle /select outlook:calendar
}
return

;------------------------------------------------
;    <CA>M -- Cycle Through Outlook Messages
;------------------------------------------------
^!M:: 
   gosub CreateOutlookMessageGroup
   GroupActivate, msg_windows, R
return

;------------------------------------------------
;    <CS>M -- Outlook New Message
;------------------------------------------------
^+M::
  Run, C:\Program Files (x86)\Microsoft Office\OFFICE14\OUTLOOK.EXE /c ipm.note
return




;******************************************************
;**
;** Mouse configuration
;**
;******************************************************
^!+`::
;   Run, main.cpl
   Run, rundll32.exe shell32.dll`, Control_RunDLL main.cpl`,`,1
   WinWait, Mouse Properties
   WinActivate
;   GuiControl, Choose, SysTabControl321, 1
;   GuiControl, ChooseString, SysTabControl321, Buttons
   Sleep, 500
   Send, ^{TAB} 
   Sleep, 50
;   GuiControl, Checkbox, Button2, 0
   ControlClick, Button2
;   ControlClick, ,Mouse Properties,OK
   Send, {ENTER}
return

;*****************************************************************************
;*****************************************************************************
;**                                                                         **
;**                       Window Control/Size Mappings                      **
;**                                                                         **
;*****************************************************************************
;*****************************************************************************

; X like move of window w/ ALT + LButton
;!LButton::
;MouseGetPos,oldmx,oldmy,mwin,mctrl
;Loop
;{
;  GetKeyState,lbutton,LButton,P
;  GetKeyState,alt,Alt,P
;  If (lbutton="U" Or alt="U")
;    Break
;  MouseGetPos,mx,my
;  WinGetPos,wx,wy,ww,wh,ahk_id %mwin%
;  wx:=wx+mx-oldmx
;  wy:=wy+my-oldmy
;  WinMove,ahk_id %mwin%,,%wx%,%wy%
;  oldmx:=mx
;  oldmy:=my
;}
;Return

; ---- Set Current Window to Always be on Top ----
^!-::
{
  WinSet,AlwaysOnTop,Toggle,A
}
return


; ---- Maximize the current window ----
; #IfWinActive ahk_class Afx:00400000:b:00010013:00000006:00140311
#IfWinActive ahk_class Afx:00400000:b:00010013:00000006:01A90B95
^!Up::gosub, handleEAaligntop
#IfWinActive
^!Up::
   WinGet,is_maximized,MinMax,A

   if is_maximized = 0
   {
      WinMaximize, A
   }
   else
   {
      WinRestore, A
   }
return

handleEAaligntop:
   WinGet,is_maximized,MinMax,A
   if is_maximized = 0
   {
      WinMaximize, A
   }
   else
   {
      Send,^!{Up}
   }
return

; #IfWinActive ahk_class Afx:00400000:b:00010013:00000006:00140311
#IfWinActive ahk_class Afx:00400000:b:00010013:00000006:01A90B95
^!Left::Send,^!{Left}
#IfWinActive
^!Left::WinMinimize, A

; #IfWinActive ahk_class Afx:00400000:b:00010013:00000006:00140311
#IfWinActive ahk_class Afx:00400000:b:00010013:00000006:01A90B95
^!Right::Send,^!{Right}
#IfWinActive
^!Right::
   monnum   := WhichMonitor()
   newwidth := GetMonitorWidth( monnum )
   newleft  := GetMonitorLeft(  monnum )

   WinMove, A,,newleft,,newwidth,
return

; --- Increase window width by 10% ---
!+Right::
   ResizeWindow(1,1.10)
return

; --- Decrease window width by 10% ---
!+Left::
   ResizeWindow(1,.90)
return

; --- Increase window height by 10% ---
!+Up::
   ResizeWindow(1.10, 1)
return

; --- Decrease window height by 10% ---
!+Down::
   ResizeWindow(.90, 1)
return

; --- Worker function to actually resize the window ---
ResizeWindow(height_delta, width_delta)
{
  WinGetPos, xpos, ypos, width, height,A

  height *= height_delta
  width  *= width_delta

  WinMove, A,,xpos, ypos, width, height
}

; --- Move window Right some ---
^!+Right::
  MoveWindow(30, 0)
return

; --- Move window left some ---
^!+Left::
  MoveWindow(-30, 0)
return

; --- Move window up some ---
^!+Up::
  MoveWindow(0, -30)
return

; --- Move window down some ---
^!+Down::
  MoveWindow(0, 30)
return

; --- Worker function to actually move the window ---
MoveWindow(xoffset, yoffset)
{
  WinGetPos, xpos, ypos, width, height,A
  xpos += xoffset
  ypos += yoffset
  WinMove, A,,xpos, ypos, width, height
}

;#IfWinActive ahk_class Afx:00400000:b:00010013:00000006:00140311
#IfWinActive ahk_class Afx:00400000:b:00010013:00000006:01A90B95
^!Down::Send,^!{Down}
#IfWinActive
; Maximize height
^!Down::
   monnum    := WhichMonitor()
   newheight := GetMonitorHeight( monnum )
   newtop    := GetMonitorTop( monnum )

   WinMove, A,,,newtop, , newheight
return

;------------------------------------------------------------
;  Use left monitor
;------------------------------------------------------------
^!1::
   UseMonitor( leftmonitor )
return

;------------------------------------------------------------
;  Use right monitor
;------------------------------------------------------------
^!2::
   UseMonitor( rightmonitor )
return

UseMonitor( monnum )
{
   UnmaximizeWindow()

   newheight := GetMonitorHeight( monnum )
   newwidth  := GetMonitorWidth( monnum )
   newleft   := GetMonitorLeft( monnum )
   newtop    := GetMonitorTop( monnum )

   WinMove, A,,newleft, newtop, newwidth, newheight

   ; Now maximize to the monitor that we're on
   WinMaximize, A
}

UseLeftHalfOfMonitor( monnum )
{
   UnmaximizeWindow()

   height := GetMonitorHeight( monnum )
   width  := GetMonitorWidth(  monnum ) / 2

   WinMove, A,,GetMonitorLeft(monnum), GetMonitorTop(monnum), width, height
}

UseRightHalfOfMonitor( monnum )
{
   UnmaximizeWindow()

   height   := GetMonitorHeight( monnum )
   width    := GetMonitorWidth(  monnum ) / 2
   leftside := GetMonitorLeft(   monnum ) + width

   WinMove, A,,leftside, GetMonitorTop(monnum), width, height
}

UseColumn( column )
{
   UnmaximizeWindow()

   height := GetMonitorHeight( 1 )
   width  := GetMonitorWidth(  1 ) / 3

   WinMove, A,,(GetMonitorLeft(monnum) + (width * (column-1)) ), GetMonitorTop(monnum), width, height
}

UseTopHalfOfMonitor( monnum )
{
   UnmaximizeWindow()

   height := GetMonitorHeight( monnum ) / 2
   width  := GetMonitorWidth(  monnum )

   WinMove, A,,GetMonitorLeft(monnum), GetMonitorTop(monnum), width, height
}

UseBottomHalfOfMonitor( monnum )
{
   UnmaximizeWindow()

   height := GetMonitorHeight( monnum ) / 2
   width  := GetMonitorWidth(  monnum )
   top    := GetMonitorTop(monnum) + height

   WinMove, A,,GetMonitorLeft(monnum), top, width, height
}

UseQuadrantOfMonitor( monnum, quadrant )
{
   UnmaximizeWindow()

   height := GetMonitorHeight( monnum ) / 2
   width  := GetMonitorWidth(  monnum ) / 2

   left   := GetMonitorLeft( monnum )
   top    := GetMonitorTop( monnum )

   if(quadrant == 1 or quadrant == 4)
   {
      left := left + width
   }

   if(quadrant == 3 or quadrant == 4)
   {
      top := top + height
   }

   WinMove, A,,left, top, width, height
}

;------------------------------------------------------------
;  Use both monitors
;------------------------------------------------------------
^!3::
   UnmaximizeWindow()

   ; figure out the height to use, smallest of the two
   mon1height := GetMonitorHeight( leftmonitor )
   mon2height := GetMonitorHeight( rightmonitor )

   if (mon1height < mon2height)
   {
      height := mon1height
   }
   else
   {
      height := mon2height
   }

   ; figure out total width
   totalwidth := GetMonitorWidth( leftmonitor ) + GetMonitorWidth( rightmonitor )

   WinMove, A,,GetMonitorLeft(leftmonitor), GetMonitorTop(leftmonitor), totalwidth, height
return

;------------------------------------------------------------
;  Use left half of left monitor
;------------------------------------------------------------
^!+1::
  UseLeftHalfOfMonitor( leftmonitor )
return

;------------------------------------------------------------
;  Use right half of left monitor
;------------------------------------------------------------
^!+2::
  UseRightHalfOfMonitor( leftmonitor )
return

;------------------------------------------------------------
;  Use top half of right monitor
;------------------------------------------------------------
^!+3::
  UseTopHalfOfMonitor( rightmonitor )
return

;------------------------------------------------------------
;  Use bottom half of right monitor
;------------------------------------------------------------
^!+4::
  UseBottomHalfOfMonitor( rightmonitor )
return

;------------------------------------------------------------
;  Switch monitors
;------------------------------------------------------------
^!+\::
  GoSub, SwitchMonitors
return

SwitchMonitors:
{
   monnum := WhichMonitor()

   WinGet,is_maximized,MinMax,A
   if(is_maximized)
   {
     if(monnum = leftmonitor)
     {
        UseMonitor(rightmonitor)
     }
     else
     {
        UseMonitor(leftmonitor)
     }

     return
   }

   WinGetPos, xpos, ypos, width, height,A

   if(monnum = rightmonitor)
   {
      xpos -= GetMonitorWidth( leftmonitor )
   }
   else
   {
      xpos += GetMonitorWidth( leftmonitor )
   }

   newmonitorheight := GetMonitorHeight( WhichMonitor() )

   WinMove, A,,xpos, ypos, width, height
   WinSet, Top
   WinActivate
   WinShow
}
return


;*****************************************************************************
;*****************************************************************************
;**                                                                         **
;**                          Program Launch Mappings                        **
;**                                                                         **
;*****************************************************************************
;*****************************************************************************


;------------------------------------------------
;                 <CAS>C -- Cycle Command Prompts
;------------------------------------------------
^!+c::
      GroupAdd, cmd_windows, ahk_class ConsoleWindowClass,,,DELLDESK
      GroupAdd, cmd_windows, ahk_class mintty,,,TODO
      GroupActivate, cmd_windows, R
return



;------------------------------------------------
;                     <AS>C -- Open Cygwin Prompt
;------------------------------------------------
!+c::
   Run, mintty -
return

!+n::
   Run, mintty screen -l -R
return

;------------------------------------------------
;                   <CS>C -- Open Command Prompts
;------------------------------------------------
^+c::
   Run, cmd, C:\
   ;Run, tcc, C:\
return


;*****************************************************************************
;*****************************************************************************
;**                                                                         **
;**                    Program Activate/Launch Mappings                     **
;**                                                                         **
;*****************************************************************************
;*****************************************************************************

;------------------------------------------------
;               <CAS>D -- Cycle DevStudio Windows
;------------------------------------------------
^!+d::
   SetTitleMatchMode, 2
   GroupAdd, msdev_windows, ahk_class wndclass_desked_gsk
   GroupAdd, msdev_windows, ahk_class TBuildMonitorForm_
   GroupAdd, msdev_windows, Microsoft Visual Studio
   GroupActivate, msdev_windows, R
return

^+d::
   if b_hasvisualstudio
   {
      Run, %visualstudiocommand%
   }
   else
   {
      MsgBox, Set visualstudiocommand variable in %config_file% to run MS Visual Studio.
   }
return


^+q::
   Run, C:\Qt\2009.05\bin\qtcreator.exe
return

^!+q::
  GroupAdd, qtcreator_windows, Qt Creator
  ; ahk_class Afx:00400000:b:00010013:00000006:01A90B95
  GroupActivate, qtcreator_windows, R
return
  

;------------------------------------------------
;  <CAS>E -- Launch/Activate Enterprise Architect
;------------------------------------------------
^!+e::
  GroupAdd, ea_windows, - EA
  ; ahk_class Afx:00400000:b:00010013:00000006:01A90B95
  GroupActivate, ea_windows, R
return

^+e::
{
   if b_hasea
   {
      Run, %eacommand%,%userprofile%/My Documents
   }
   else
   {
      MsgBox, Set eacommand in %config_file% to run Enterprise Architect
   }
}
return

;------------------------------------------------
;               <CAS>F -- Launch/Activate Firefox
;------------------------------------------------
^!+f::
  GroupAdd, firefox_windows, ahk_class MozillaUIWindowClass
  GroupActivate, firefox_windows, R
  WinShow
return

;------------------------------------------------
;          <CAS>G -- Cycle Google Chrome Windows
;------------------------------------------------
^!+g::
   SetTitleMatchMode, 2
   GroupAdd, chrome_windows, ahk_class Chrome_WidgetWin_1,,,Grooveshark
   GroupActivate, chrome_windows, R
return

^+g::
  if b_haschrome
  {
     Run, %chromecommand%
  }
  else
  {
     MsgBox, Set chromecommand in %config_file% to run Chrome
  }
return

;------------------------------------------------
;               <CAS>I -- Launch/Activate Safari
;------------------------------------------------
^!+i::
   GroupAdd, safari_windows, ahk_class {1C03B488-D53B-4a81-97F8-754559640193}
   GroupActivate, safari_windows, R
   WinShow
return

;------------------------------------------------
;               <C S>S -- Launch SCM Tool
;------------------------------------------------
^+s::
  if b_hasscm
  {
     Run, %scmcommand%
  }
  else
  {
     MsgBox, Set scmcommand in %config_file% to launch SCM tool.
  }
return

^!s::
{
   SetTitleMatchMode, 2
   GroupAdd, scm_windows, Perforce P4V
   GroupAdd, scm_windows, Perforce Password Required
   GroupActivate, scm_windows, R
}
return


;------------------------------------------------
;         <CS>U -- Launch Ubuntu VM
;------------------------------------------------
^+u::
   Run, c:\program files\Sun\xVM VirtualBox\VBoxManage.exe startvm Kubuntu
return

;------------------------------------------------
;         <CAS>U -- Activate Ubuntu VM
;------------------------------------------------
^!+u::
  SetTitleMatchMode,2
  IfWinExist, Sun VirtualBox
  {
    WinActivate
  }
return



;--------------------------------------------
;  Switch between Putty Windows (_H_ome)
;--------------------------------------------
^!+h::
  GroupAdd, putty_windows, ahk_class PuTTY
  GroupActivate, putty_windows, R
return

;--------------------------------------------
;  Log into home computer
;--------------------------------------------
^+h::
  Run, C:\utils\Putty @Home
return

;------------------------------------------------
; Launch KeePass
;------------------------------------------------
^+k::
  Run, c:\KeePass\KeePass.exe, C:\KP
return

;------------------------------------------------
;                           <CAS>V -- launch gVim
;------------------------------------------------
^!+v::
SetTitleMatchMode,2
  GroupAdd, gvim_windows, ahk_class Vim 
  GroupActivate, gvim_windows, R
return

; Explicitly Launch gVim
^+v::
  EnvSet, ComSpec, %SystemRoot%\system32\cmd.exe
  Run, %vimcommand%, C:\
return

;------------------------------------------------
;                          <CS>W - Launch MS Word
;------------------------------------------------
^+w::
  if b_hasmsword
  {
     Run, %mswordcommand%
  }
  else
  {
     MsgBox, Set mswordcommand in %config_file% to run MS Word.
  }
return

;------------------------------------------------
;                         <CS>E - Launch MS Excel
;------------------------------------------------
^+x::
  if b_hasexcel
  {
     Run, %excelcommand%
  }
  else
  {
     MsgBox, Set excelcommand in %config_file% to run MS Excel.
  }
return

!+x::
  ;Run, C:\cygwin\usr\X11R6\bin\startxwin.bat
  Run, XWin -multiwindow -clipboard -silent-dup-error
return

;----------------------------------------------------
;                 <CAS><ESC> Launch Process Explorer
;----------------------------------------------------
^!+Esc::
  Run, c:\utils\procexp\procexp.exe
return


;*****************************************************************************
;*****************************************************************************
;**                                                                         **
;**                        Window Cycling Mappings                          **
;**                                                                         **
;*****************************************************************************
;*****************************************************************************

;------------------------------------------------
;          <CAS>W - Cycle Between MS Word Windows
;------------------------------------------------
^!+w::
  SetTitleMatchMode,2
  GroupAdd, MSWord, Microsoft Word
  GroupAdd, MSWord, LibreOffice Writer
  GroupActivate, MSWord, R
return

;------------------------------------------------
;         <CAS>E - Cycle Between MS Excel Windows
;------------------------------------------------
^!+x::
   SetTitleMatchMode, 2
   GroupAdd, MSExcel, ahk_class XLMAIN
   GroupAdd, MSExcel, LibreOffice Calc
   GroupActivate, MSExcel, R
return

;------------------------------------------------
;   <CAS>R - Cycle Between Acrobat Reader Windows
;------------------------------------------------
^!+r::
  SetTitleMatchMode,2
  GroupAdd, AcroReader, Adobe Reader
  GroupAdd, AcroReader, PDF-XChange Viewer
  GroupActivate, AcroReader, R
return



;*****************************************************************************
;*****************************************************************************
;**                                                                         **
;**                          Multimedia Key Mappings                        **
;**                                                                         **
;*****************************************************************************
;*****************************************************************************

;------------------------------------------------
;                       <AS>L - Listen for Client
;------------------------------------------------
!+l::
   if(!b_isserver)
   {
      TrayTip, Rex's Macros, AHK script not configured as a server!, 5, 1
      return
   }

   TrayTip, Rex's Macros, Waiting for remote connection..., 5, 1

   ; Network_Address = 127.0.0.1
   Network_Address = 10.102.1.18
   ; socket := PrepareForIncomingConnection(Network_Address, commport)
   socket := PrepareForIncomingConnection(ServerIP, commport)
    
   ; Find this script's main window:
   Process, Exist  ; This sets ErrorLevel to this script's PID (it's done this way to support compiled scripts).
   DetectHiddenWindows On
   ScriptMainWindowId := WinExist("ahk_class AutoHotkey ahk_pid " . ErrorLevel)
   DetectHiddenWindows Off

   ; When the OS notifies the script that there is incoming data waiting to be received,
   ; the following causes a function to be launched to read the data:
   OnMessage(WinMessageID, "ReceiveData")

   ; Set up the connection to notify this script via message whenever new data has arrived.
   ; This avoids the need to poll the connection and thus cuts down on resource usage.
   FD_READ = 1     ; Received when data is available to be read.
   FD_CLOSE = 32   ; Received when connection has been closed.
   FD_CONNECT = 20 ; Recieved when connection has been made.
   if DllCall("Ws2_32\WSAAsyncSelect", "UInt", socket, "UInt", ScriptMainWindowId, "UInt", WinMessageID, "Int", FD_READ|FD_CLOSE|FD_CONNECT)
   {
       MsgBox % "WSAAsyncSelect() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
       return
   }

   tries = 50
   Loop ; Wait for incomming connections
   {
      ; accept requests that are in the pipeline of the socket   
      conectioncheck := DllCall("Ws2_32\accept", "UInt", socket, "UInt", &SocketAddress, "Int", SizeOfSocketAddress)

      ; Ws2_22/accept returns the new Connection-Socket if a connection request was in the pipeline
      ; on failure it returns an negative value
       if conectioncheck > 1
       {
          TrayTip, Rex's Macros, Incoming connection accepted, 5, 1
          break   
      }

      sleep 500 ; wait half second then accept again

      tries := tries - 1

      ;tooltip, Waiting for Remote AHK connection: %tries%

      if(tries = 0)
      {
         break
      }
   }   

   if(tries = 0)
   {
      TrayTip, Rex's Macros, No remote connection made..., 5, 1
   }

return

;------------------------------------------------
;                             <CAS>k -- Volume Up
;------------------------------------------------
^!+k::
  if(b_isclient)
  {
    SendData(socket,"volup")
  }
  else
  {
    Gosub, volup
  }
return

^!+WheelUp::
  Gosub, volup
return

volup:
; Send, {VOLUME_UP}
   SoundSet, +2
   Gosub, vol_OSD
return

;------------------------------------------------
;                           <CAS>j -- Volume Down
;------------------------------------------------
^!+j::
  if(b_isclient)
  {
    SendData(socket,"voldown")
  }
  else
  {
    Gosub, voldown
  }
return

^!+WheelDown::
  Gosub, voldown
return

voldown:
;Send, {VOLUME_DOWN}
   SoundSet, -2
   Gosub, vol_OSD
return

;------------------------------------------------
;                                  <CAS>M -- Mute
;------------------------------------------------
^!+m::
  if(b_isclient)
  {
     SendData(socket,"mute")
  }
  else
  {
    Gosub, mute
  }
return

mute:
     ;Send, {VOLUME_MUTE}
     SoundSet, +1, , mute
     Gosub, vol_OSD
return

;------------------------------------------------
;                            <CAS>p -- Play/Pause
;------------------------------------------------
^!+p::
  if(b_isclient)
  {
    SendData(socket,"playpause")
  }
  else
  {
     gosub, playpause
  }
return

;**************************************************************************
playpause: ;{{{
   SetTitleMatchMode,2
   IfWinExist, Grooveshark
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      ; Cache the mouse position to restore it later
      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate

      ; Click the play/pause button
      Sleep, 200
      MouseClick, Left,25,975,1,0

      ; Restore the window show/hide state
      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      ; Restore the mouse position
      MouseMove, xmousepos, ymousepos
   }
   IfWinExist, Finetune Desktop
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      ; Cache the mouse position to restore it later
      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate

      ; Click the play/pause button
      MouseClick, Left,35,270,1,0

      ; Restore the window show/hide state
      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      ; Restore the mouse position
      MouseMove, xmousepos, ymousepos

   }
   else IfWinExist, Pandora Radio
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate
   ;   MouseClick, Left,10,120,1,0
      MouseClick, Left,786,161,1,0

      ; Give it time to settle in
      Sleep, 100

      Send, {SPACE}

      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      MouseMove, xmousepos, ymousepos
   }
   else IfWinExist, Rhapsody Player
   {
      ; Get current state
      MouseGetPos, xmousepos, ymousepos
      WinGet, active_id, ID, A
      WinGet, is_minimized, MinMax, Rhapsody Player

      WinActivate

      WinGetPos,,,rhapwidth,,A
      buttonxpos:=rhapwidth-180
      activatexpos:=rhapwidth-180

      ; Activate the Flash control
      Send {Click %activatexpos%, 95}

      ; Activate Rhapsody player and click play/pause button
      Send {Click %buttonxpos%, 245}

      ; Restore previous state
      if(is_minimized=-1)
      {
         WinMinimize
      }
      WinActivate ahk_id %active_id%
      MouseMove, xmousepos, ymousepos
   }
   else if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
   {
;      ControlClick, TEST_WIN32WND92

      minimize=1

      IfWinActive
      {
         minimize=0
      }

      WinRestore
      WinActivate

      Send, {MEDIA_PLAY_PAUSE}

      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }
   }
   else
   {
      Send, {MEDIA_PLAY_PAUSE}
   }
return ; }}}
;**************************************************************************

;------------------------------------------------
;                                  <CAS>s -- Stop
;------------------------------------------------
^!+s::
  if(b_isclient)
  {
    SendData(socket,"stop")
  }
  else
  {
     gosub, stop
  }
return

stop:
   if WinExist("ahk_class TEST_WIN32WND")  ; Rhapsody
   {
   ;   ControlClick, TEST_WIN32WND135
   ;   ControlClick, TEST_WIN32WND11
   ;   ControlClick, TEST_WIN32WND188
      ControlClick, TEST_WIN32WND94
   }
   else
   {
      Send, {MEDIA_STOP}
   }
return

;------------------------------------------------
;                           <CAS> > -- Next track
;------------------------------------------------
^!+>::
  if(b_isclient)
  {
    SendData(socket,"nexttrack")
  }
  else
  {
     gosub, nexttrack
  }
return


nexttrack:
   SetTitleMatchMode,2
   IfWinExist, Grooveshark
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      ; Cache the mouse position to restore it later
      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate

      ; Click the next track button
      Sleep, 200
      MouseClick, Left,90,975,1,0

      ; Restore the window show/hide state
      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      ; Restore the mouse position
      MouseMove, xmousepos, ymousepos
   }
   IfWinExist, Finetune Desktop
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate

      Sleep, 100

      MouseClick, Left,210,180,1,0

      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      MouseMove, xmousepos, ymousepos

   }
   else if WinExist("Pandora Radio")
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate

      MouseClick, Left,60,390,1,0
      Send, {RIGHT}

      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      MouseMove, xmousepos, ymousepos
   }
   else IfWinExist, Rhapsody Player
   {
;      ; Get current state
;      MouseGetPos, xmousepos, ymousepos
;      WinGet, active_id, ID, A
;      WinGet, is_minimized, MinMax, Rhapsody Player
;
;      ; Activate Rhapsody player and click next track button
;      WinActivate
;      Send {Click 100, 175}
;
;      ; Restore previous state
;      if(is_minimized=-1)
;      {
;         WinMinimize
;      }
;      WinActivate ahk_id %active_id%
;      MouseMove, xmousepos, ymousepos
      ; Get current state
      MouseGetPos, xmousepos, ymousepos
      WinGet, active_id, ID, A
      WinGet, is_minimized, MinMax, Rhapsody Player

      WinActivate

      WinGetPos,,,rhapwidth,,A
      buttonxpos:=rhapwidth-70
      activatexpos:=rhapwidth-180

      ; Activate the Flash control
      Send {Click %activatexpos%, 95}

      ; Activate Rhapsody player and click next track button
      Send {Click %buttonxpos%, 245}

      ; Restore previous state
      if(is_minimized=-1)
      {
         WinMinimize
      }
      WinActivate ahk_id %active_id%
      MouseMove, xmousepos, ymousepos
   }
   else if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
   {
      ;ControlClick, TEST_WIN32WND90

      minimize=1

      IfWinActive
      {
         minimize=0
      }

;      WinRestore
      WinActivate

      Send, {MEDIA_NEXT}

      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }
   }
   else
   {
      Send, {MEDIA_NEXT}
   }
return



;------------------------------------------------
;                       <CAS> < -- Previous track
;------------------------------------------------
^!+<::
  if(b_isclient)
  {
    SendData(socket,"prevtrack")
  }
  else
  {
     gosub, prevtrack
  }
return

prevtrack:
   SetTitleMatchMode,2
   IfWinExist, Grooveshark
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      ; Cache the mouse position to restore it later
      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate

      ; Click the next track button
      Sleep, 200
      MouseClick, Left,60,975,1,0

      ; Restore the window show/hide state
      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      ; Restore the mouse position
      MouseMove, xmousepos, ymousepos
   }
   IfWinExist, Finetune Desktop
   {
      minimize=1

      IfWinActive
      {
         minimize=0
      }

      MouseGetPos, xmousepos, ymousepos
      WinRestore
      WinActivate

      Sleep, 100

      MouseClick, Left,40,180,1,0

      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }

      MouseMove, xmousepos, ymousepos

   }
   else IfWinExist, Rhapsody Player
   {
      ; Get current state
      MouseGetPos, xmousepos, ymousepos
      WinGet, active_id, ID, A
      WinGet, is_minimized, MinMax, Rhapsody Player

      WinActivate

      WinGetPos,,,rhapwidth,,A
      buttonxpos:=rhapwidth-265
      activatexpos:=rhapwidth-180

      ; Activate the Flash control
      Send {Click %activatepos%, 95}

      ; Activate Rhapsody player and click next track button
      Send {Click %buttonxpos%, 245}

      ; Restore previous state
      if(is_minimized=-1)
      {
         WinMinimize
      }
      WinActivate ahk_id %active_id%
      MouseMove, xmousepos, ymousepos
   }
   else if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
   {
      ;ControlClick, TEST_WIN32WND91

      minimize=1

      IfWinActive
      {
         minimize=0
      }

      WinRestore
      WinActivate

      Send, {MEDIA_PREV}

      if minimize>0
      {
         Send, {ALTDOWN}{TAB}{ALTUP}
         WinMinimize
      }
   }
   else
   {
      Send, {MEDIA_PREV}
   }
return


;------------------------------------------------
; Rhapsody Ratings, 1-5 <A-S><1>..<5>
;------------------------------------------------
!+0::
if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
{
   ControlClick, TEST_WIN32WND77         ; Don't play!
}

return
!+1::
if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
{
   ControlClick, TEST_WIN32WND76         ; One star
}
return

!+2::
if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
{
   ControlClick, TEST_WIN32WND75         ; Two stars
}
return

!+3::
if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
{
   ControlClick, TEST_WIN32WND74         ; Three stars

}
return

!+4::
if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
{
   ControlClick, TEST_WIN32WND73         ; Four stars
}
return

!+5::
if WinExist("ahk_class TEST_WIN32WND")   ; Rhapsody
{
   ControlClick, TEST_WIN32WND72         ; Five stars
}
return

;------------------------------------------------
;                   <CA>j -- Eject/Insert CD Tray
;------------------------------------------------
^!j::
Drive, Eject
if A_TimeSinceThisHotkey < 1000
     Drive, Eject,, 1
return

;------------------------------------------------
;                   <CS>j -- Eject USB Devices
;------------------------------------------------
^+j::
{
   Run, rundll32.exe shell32.dll`, Control_RunDLL hotplug.dll`
   return
}

;------------------------------------------------
;          <CAS> ? -- Launch/Activate MediaPlayer
;------------------------------------------------
^!+?::
   SetTitleMatchMode,2
   IfWinExist, Grooveshark
   {
      WinRestore
      WinActivate
   }
   IfWinExist, Finetune Desktop
   {
      WinRestore
      WinActivate
   }
   else IfWinExist, Pandora Radio
   {
      MouseGetPos, xmousepos, ymousepos

      WinRestore
      WinActivate

      MouseClick, Left,60,390,1,0
      MouseMove, xmousepos, ymousepos
   }
   else IfWinExist, Rhapsody Player
   {
      WinRestore
      WinActivate
   }
   else if WinExist(ahk_class %RHAPSODY_WINCLASS%)   ; Rhapsody
   {
     WinActivate
   }
   else if WinExist("ahk_class BaseWindow_RootWnd")
   {
      WinActivate
   }
   else if WinExist("ahk_class WMPlayerApp")
   {
     WinActivate
   }
   else IfWinExist, Player Window
   {
     WinActivate
   }
   else
   {
     ; Run, C:\Program Files\Windows Media Player\wmplayer.exe /prefetch:1  
     ; Run, C:\Program Files\Rhapsody\Rhapsody.exe,C:\
     Run, %defaultmediaplayercommand%
   }
return

; Open AC97 equalizer
^!+z::
   Run, rundll32.exe shell32.dll`, Control_RunDLL alsndmgr.cpl`,`,1
return

; Equalize to live
^!z::
   Run, rundll32.exe shell32.dll`, Control_RunDLL alsndmgr.cpl`,`,1
   WinWait, AC97 Audio Configuration
   WinActivate
   ControlClick, Button12 ; Live
   Send, {ESC}
return

; Equalize to Rap
^+z::
   Run, rundll32.exe shell32.dll`, Control_RunDLL alsndmgr.cpl`,`,1
   WinWait, AC97 Audio Configuration
   WinActivate
   ControlClick, Button2 ; Load
   WinWait, Load Preset
   WinActivate
   Control, ChooseString, Rap
   Send, {Enter}
   WinWait, AC97 Audio Configuration
   WinActivate
   Send, {Esc}
return

; Equalize to Rock
!+z::
   Run, rundll32.exe shell32.dll`, Control_RunDLL alsndmgr.cpl`,`,1
   WinWait, AC97 Audio Configuration
   WinActivate
   ControlClick, Button16 ; Rock
   Send, {ESC}
return


;--------------------------------------------------
;                    <CAS>+ -- Pandora, I LIKE IT!
;--------------------------------------------------
^!++::
; if WinExist("Pandora Desktop Player")
IfWinExist, Pandora Radio
{
   minimize=1

   IfWinActive
   {
      minimize=0
   }

   MouseGetPos, xmousepos, ymousepos
   WinRestore
   WinActivate
;   MouseClick, Left,10,120,1,0
   MouseClick, Left,60,390,1,0
   Send, =

   if minimize>0
   {
      Send, {ALTDOWN}{TAB}{ALTUP}
      WinMinimize
   }

   MouseMove, xmousepos, ymousepos
}
return

;--------------------------------------------------
;                    <CAS>- -- Pandora, I HATE IT!
;--------------------------------------------------
^!+-::
; if WinExist("Pandora Desktop Player")
IfWinExist, Pandora Radio
{
   minimize=1

   IfWinActive
   {
      minimize=0
   }

   MouseGetPos, xmousepos, ymousepos
   WinRestore
   WinActivate
;  MouseClick, Left,10,120,1,0
   MouseClick, Left,60,390,1,0
   Send, -

   if minimize>0
   {
      Send, {ALTDOWN}{TAB}{ALTUP}
      WinMinimize
   }

   MouseMove, xmousepos, ymousepos
}
return


;--------------------------------------------------
; ---------- On Screen Display --------------------
;--------------------------------------------------

vol_OSD:
  if b_useVolumeOSD
  {
     IfWinNotExist, vol_Master
     {
        ; Calculate position here in case screen resolution changes
        ; while the script is running:
        if vol_PosY < 0
        {
           ; Create the Wave bar just above the Master bar:
           WinGetPos, , vol_Wave_Posy, , , vol_Wave
           vol_Wave_Posy -= %vol_Thick%
           Progress, 1:B ZH12 ZX0 ZY0 W150 CBRed CWSilver, ,, vol_Master
        }
        else
        {    
           Progress, 1:B ZH12 ZX0 ZY0 W150 CBRed CWSilver, ,, vol_Master
        }
     }

     SoundGet, vol_Master, Master
     Progress, 1:%vol_Master%

     SoundGet, master_mute, , mute
     if master_mute != Off
     {
        Progress, 1:0
     }

     SetTimer, vol_OSD_Off, 800
  }
return


vol_OSD_Off:
  SetTimer, vol_OSD_Off, off
  Progress, 1:Off
return


;-------------------------------------------------------------
; Helper Functions
;-------------------------------------------------------------
WhichMonitor()
{
   global leftmonitor
   global rightmonitor

   middle := GetMonitorLeft(rightmonitor)

   WinGetPos, xpos, , , ,A

   ; HACK: account for the edges of the monitor when the window is maximized
   xpos += 10 

   if(xpos < middle)
   {
      return leftmonitor
   }
   else
   {
      return rightmonitor
   }
}

GetMonitorWidth(monnum)
{
   SysGet, MonSz, MonitorWorkArea, %monnum%

   newwidth := Abs(MonSzRight - MonSzLeft)

   return newwidth
}

GetMonitorHeight(monnum)
{
   SysGet, MonSz, MonitorWorkArea, %monnum%

   newheight := Abs(MonSzTop - MonSzBottom)

   return newheight
}

GetMonitorTop(monnum)
{
   SysGet, MonSz, MonitorWorkArea, %monnum%
   return MonSzTop
}

GetMonitorLeft(monnum)
{
   SysGet, MonSz, MonitorWorkArea, %monnum%
   return MonSzLeft
}


IsCloseTo(constant, variable)
{
  vp := variable + 5
  vm := variable - 5

  if( constant < vp )
  {
     if(constant > vm)
     {
        return 1
     }
  }

  return 0
}

Min(a, b)
{
   if(a < b)
   {
      return a
   }
   else
   {
      return b
   }
}

Max(a, b)
{
   if(a > b)
   {
      return a
   }
   else
   {
      return b
   }
}

Abs(v)
{
   if(v < 0)
   {
      v *= -1
   }

   return v
}

UnmaximizeWindow()
{
   ; If the current window is currently maximize it, unmaximize it.
   WinGet,is_maximized,MinMax,A
   if(is_maximized)
   {
      WinRestore, A
   }
}



; -----------------------------------------------------
;             Paste Plain Text into any app
; -----------------------------------------------------

; This function is called the first time that you hit ^V, which sets a
; timer for the second press
PASTEONCE:
pastecounter+=1
SetTimer,PASTETWICE,%pasteagaintimeout%
Return

;  
PASTETWICE:
SetTimer,PASTETWICE,Off
wholeclipboard:=ClipboardAll
If pastecounter>1
{
   Clipboard=%Clipboard%
}

pastecounter=0
Send,^v
Clipboard:=wholeclipboard
Return


;**************************************************************************
; Menu Functions                                                 {{{
;-------------------------------------------------------------------
;-------------------------------------------------------------------

; ---------------------------------------------------------------------------- 
;    Multi Launcher menu items
; ---------------------------------------------------------------------------- 
MultiLauncherFunction:
   if A_ThisMenuItemPos = 1
      {
         if b_hasfirefox
         {
            Run, %firefoxcommand%, C:\
         }
         else
         {
            MsgBox, Set firefoxcommand variable in %config_file% to run Firefox
            Run, %fallbackbrowser%
         }
      }
   if A_ThisMenuItemPos = 2
   {
      SetTitleMatchMode, 3

      IfWinExist, Calculator
      {
         WinActivate
      }
      else
      {
         Run, calc
      }
   }
   ; -- After this is the run clipboard contents item
return

MiscCommandsFunction:
   if A_ThisMenuItemPos = 1 ; Close all explorer windows
   {
      GroupAdd, explorer_windows, ahk_class CabinetWClass
      GroupAdd, explorer_windows, ahk_class ExploreWClass
      GroupClose, explorer_windows, A
   }
   if A_ThisMenuItemPos = 2 ; Close all e-mail windows
   {
      gosub CreateOutlookMessageGroup
      GroupClose, msg_windows, A
   }
   if A_ThisMenuItemPos = 3 ; Update Doxyments
   {
      IfWinExist, Doxygen GUI frontend
      {
         WinActivate
         ControlClick, QWidget13
      }
   }
return

; ---------------------------------------------------------------------------- 
;    Clipboard Search functions
; ---------------------------------------------------------------------------- 
HandleClipboard:
   if A_ThisMenuItemPos = 1
     Run, IEXPLORE.EXE "%Clipboard%"
   if A_ThisMenuItemPos = 2
     Run, http://www.google.com/search?q=%Clipboard%
   if A_ThisMenuItemPos = 3
     Run, http://127.0.0.1:4664/search?q=%Clipboard%&flags=68&num=10&s=R04pOTnDutX9WBY8-LWb4bxkPAU
   if A_ThisMenuItemPos = 4
     Run, http://en.wikipedia.org/wiki/Special:Search?search=%Clipboard%
   if A_ThisMenuItemPos = 5
     Run, http://m-w.com/dictionary/%Clipboard%
   if A_ThisMenuItemPos = 7
     AddEARequirement(0)
   if A_ThisMenuItemPos = 8
     AddEARequirement(1)
return

;                                                                }}}
;**************************************************************************


RunClipboard:
  Run, %Clipboard%
return

SendClipboard:
  cbstring = cbð%clipboard%
  SendData(socket, cbstring)
return


;**************************************************************************
;    Favorites Menu                                              {{{
;-------------------------------------------------------------------
;-------------------------------------------------------------------

;----Open the selected favorite
f_OpenFavorite:
; Fetch the array element that corresponds to the selected menu item:
StringTrimLeft, f_path, f_path%A_ThisMenuItemPos%, 0
if f_path =
	return
if f_class = #32770    ; It's a dialog.
{
	if f_Edit1Pos <>   ; And it has an Edit1 control.
	{
		; Activate the window so that if the user is middle-clicking
		; outside the dialog, subsequent clicks will also work:
		WinActivate ahk_id %f_window_id%
		; Retrieve any filename that might already be in the field so
		; that it can be restored after the switch to the new folder:
		ControlGetText, f_text, Edit1, ahk_id %f_window_id%
		ControlSetText, Edit1, %f_path%, ahk_id %f_window_id%
		ControlSend, Edit1, {Enter}, ahk_id %f_window_id%
		Sleep, 100  ; It needs extra time on some dialogs or in some cases.
		ControlSetText, Edit1, %f_text%, ahk_id %f_window_id%
		return
	}
	; else fall through to the bottom of the subroutine to take standard action.
}
else if f_class in ExploreWClass,CabinetWClass  ; In Explorer, switch folders.
{
	if f_Edit1Pos <>   ; And it has an Edit1 control.
	{
		ControlSetText, Edit1, %f_path%, ahk_id %f_window_id%
		; Tekl reported the following: "If I want to change to Folder L:\folder
		; then the addressbar shows http://www.L:\folder.com. To solve this,
		; I added a {right} before {Enter}":
		ControlSend, Edit1, {Right}{Enter}, ahk_id %f_window_id%
		return
	}
	; else fall through to the bottom of the subroutine to take standard action.
}
else if f_class = ConsoleWindowClass ; In a console window, CD to that directory
{
	WinActivate, ahk_id %f_window_id% ; Because sometimes the mclick deactivates it.
	SetKeyDelay, 0  ; This will be in effect only for the duration of this thread.
	IfInString, f_path, :  ; It contains a drive letter
	{
		StringLeft, f_path_drive, f_path, 1
		Send %f_path_drive%:{enter}
	}
	Send, cd %f_path%{Enter}
	return
}

Run, Explorer %f_path%  ; Might work on more systems without double quotes.
return


;----Display the menu
f_DisplayMenu:
; These first few variables are set here and used by f_OpenFavorite:
WinGet, f_window_id, ID, A
WinGetClass, f_class, ahk_id %f_window_id%
if f_class in #32770,ExploreWClass,CabinetWClass  ; Dialog or Explorer.
	ControlGetPos, f_Edit1Pos,,,, Edit1, ahk_id %f_window_id%

if f_class in #32770,ExploreWClass,CabinetWClass  ; Dialog or Explorer.
{
   if f_Edit1Pos =  ; The control doesn't exist, so don't display the menu
      return
}
; else if f_class <> ConsoleWindowClass
   ; return ; Since it's some other window type, don't display menu.

; Otherwise, the menu should be presented for this type of window:
Menu, Favorites, show
return

;                                                                }}}
;**************************************************************************
; Client Server support                                          {{{
;-------------------------------------------------------------------
;-------------------------------------------------------------------
PrepareForIncomingConnection(IPAddress, Port)
; This can connect to most types of TCP servers, not just Network.
; Returns -1 (INVALID_SOCKET) upon failure or the socket ID upon success.
{
    VarSetCapacity(wsaData, 32)  ; The struct is only about 14 in size, so 32 is conservative.
    result := DllCall("Ws2_32\WSAStartup", "UShort", 0x0002, "UInt", &wsaData) ; Request Winsock 2.0 (0x0002)
    ; Since WSAStartup() will likely be the first Winsock function called by this script,
    ; check ErrorLevel to see if the OS has Winsock 2.0 available:
    if ErrorLevel
    {
        MsgBox WSAStartup() could not be called due to error %ErrorLevel%. Winsock 2.0 or higher is required.
        return -1
    }
    if result  ; Non-zero, which means it failed (most Winsock functions return 0 upon success).
    {
        MsgBox % "WSAStartup() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
        return -1
    }

    AF_INET = 2
    SOCK_STREAM = 1
    IPPROTO_TCP = 6
    socket := DllCall("Ws2_32\socket", "Int", AF_INET, "Int", SOCK_STREAM, "Int", IPPROTO_TCP)
    if socket = -1
    {
        MsgBox % "socket() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
        return -1
    }

    ; Prepare for connection:
    SizeOfSocketAddress = 16
    VarSetCapacity(SocketAddress, SizeOfSocketAddress)
    InsertInteger(2, SocketAddress, 0, AF_INET)   ; sin_family
    InsertInteger(DllCall("Ws2_32\htons", "UShort", Port), SocketAddress, 2, 2)   ; sin_port
    InsertInteger(DllCall("Ws2_32\inet_addr", "Str", IPAddress), SocketAddress, 4, 4)   ; sin_addr.s_addr

    ; Bind to socket:
    if DllCall("Ws2_32\bind", "UInt", socket, "UInt", &SocketAddress, "Int", SizeOfSocketAddress)
    {
        MsgBox % "bind() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError") . "?"
        return -1
    }
    if DllCall("Ws2_32\listen", "UInt", socket, "UInt", "SOMAXCONN")
    {
        MsgBox % "LISTEN() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError") . "?"
        return -1
    }
   
    return socket  ; Indicate success by returning a valid socket ID rather than -1. 
}


ReceiveData(wParam, lParam)
; By means of OnMessage(), this function has been set up to be called automatically whenever new data
; arrives on the connection. 
{
   ; global ShowRecieved
   socket := wParam
   ReceivedDataSize = 4096  ; Large in case a lot of data gets buffered due to delay in processing previous data.

   ; This loop solves the issue of the notification message being discarded due to thread-already-running.
   Loop  
   {
     VarSetCapacity(ReceivedData, ReceivedDataSize, 0)  ; 0 for last param terminates string for use with recv().
     ReceivedDataLength := DllCall("Ws2_32\recv", "UInt", socket, "Str", ReceivedData, "Int", ReceivedDataSize, "Int", 0)

     if ReceivedDataLength = 0  ; The connection was gracefully closed,
     {
        TrayTip, Rex's Macros, Remote AHK command client disconnected..., 5, 1
        DllCall("Ws2_32\WSACleanup")
        return
        ;    ExitApp  ; The OnExit routine will call WSACleanup() for us.
     }

     if ReceivedDataLength = -1
     {
        WinsockError := DllCall("Ws2_32\WSAGetLastError")
        if WinsockError = 10035  ; WSAEWOULDBLOCK, which means "no more data to be read".
        {
          return 1
        }

        if WinsockError <> 10054 ; WSAECONNRESET, which happens when Network closes via system shutdown/logoff.
        {
          ; Since it's an unexpected error, report it.  Also exit to avoid infinite loop.
          TrayTip, Rex's Macros, recv() indicated Winsock error %WinsockError%
          DllCall("Ws2_32\WSACleanup")
          ; ExitApp  ; The OnExit routine will call WSACleanup() for us.
          return
        }
     }

      Loop, parse, ReceivedData, `n, `r
      {
        ; ShowRecieved = %ShowRecieved%%A_LoopField%
        ShowRecieved = %A_LoopField%

        StringSplit, CommandsArray, ShowRecieved, `;
        Loop, %CommandsArray0%
        {
          command := CommandsArray%a_index%

          IfInString, command, volup
          {
            gosub, volup
          }
          else IfInString, command, voldown
          {
            gosub, voldown
          }
          else IfInString, command, mute
          {
            Gosub, mute
          }
          else IfInString, command, playpause
          {
            gosub, playpause
          }
          else IfInString, command, nexttrack
          {
            gosub, nexttrack
          }
          else IfInString, command, prevtrack
          {
            gosub, prevtrack
          }
          else IfInString, command, stop
          {
            gosub, stop
          }
          else IfInString, command, cb
          {
            StringSplit, cbarray, command, `ð
            Clipboard := cbarray2
            TrayTip, Rex's Macros, Received new clipboard contents, 5, 1
          }
          else
          {
             MsgBox, Ooops! %command%
          }
        }
        
        ShowReceived = ""
      }
    }

    return 1  ; Tell the program that no further processing of this message is needed.
}

SendData(wParam,SendData)
{
   socket := wParam
   SendDataSize := VarSetCapacity(SendData)
   SendDataSize += 1
   sendret := DllCall("Ws2_32\send", "UInt", socket, "Str", SendData, "Int", SendDatasize, "Int", 0)
} 

ConnectToAddress(IPAddress, Port)
; Returns -1 (INVALID_SOCKET) upon failure or the socket ID upon success.
{
    VarSetCapacity(wsaData, 32)  ; The struct is only about 14 in size, so 32 is conservative.
    result := DllCall("Ws2_32\WSAStartup", "UShort", 0x0002, "UInt", &wsaData) ; Request Winsock 2.0 (0x0002)
    ; Since WSAStartup() will likely be the first Winsock function called by this script,
    ; check ErrorLevel to see if the OS has Winsock 2.0 available:
    if ErrorLevel
    {
        MsgBox WSAStartup() could not be called due to error %ErrorLevel%. Winsock 2.0 or higher is required.
        return -1
    }
    if result  ; Non-zero, which means it failed (most Winsock functions return 0 upon success).
    {
        MsgBox % "WSAStartup() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
        return -1
    }

    AF_INET = 2
    SOCK_STREAM = 1
    IPPROTO_TCP = 6
    socket := DllCall("Ws2_32\socket", "Int", AF_INET, "Int", SOCK_STREAM, "Int", IPPROTO_TCP)
    if socket = -1
    {
        MsgBox % "socket() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError")
        return -1
    }

    ; Prepare for connection:
    SizeOfSocketAddress = 16
    VarSetCapacity(SocketAddress, SizeOfSocketAddress)
    InsertInteger(2, SocketAddress, 0, AF_INET)   ; sin_family
    InsertInteger(DllCall("Ws2_32\htons", "UShort", Port), SocketAddress, 2, 2)   ; sin_port
    InsertInteger(DllCall("Ws2_32\inet_addr", "Str", IPAddress), SocketAddress, 4, 4)   ; sin_addr.s_addr

    ; Attempt connection:
    if DllCall("Ws2_32\connect", "UInt", socket, "UInt", &SocketAddress, "Int", SizeOfSocketAddress)
    {
        MsgBox % "connect() indicated Winsock error " . DllCall("Ws2_32\WSAGetLastError") . "?"
        return -1
    }
    return socket  ; Indicate success by returning a valid socket ID rather than -1.
}

InsertInteger(pInteger, ByRef pDest, pOffset = 0, pSize = 4)
; The caller must ensure that pDest has sufficient capacity.  To preserve any existing contents in pDest,
; only pSize number of bytes starting at pOffset are altered in it.
{
    Loop %pSize%  ; Copy each byte in the integer into the structure as raw binary data.
        DllCall("RtlFillMemory", "UInt", &pDest + pOffset + A_Index-1, "UInt", 1, "UChar", pInteger >> 8*(A_Index-1) & 0xFF)
}

ShutdownClientServer:
; MSDN: "Any sockets open when WSACleanup is called are reset and
;        automatically deallocated as if closesocket was called."
DllCall("Ws2_32\WSACleanup")
ExitApp
;                                                                }}}
;**************************************************************************

AddEARequirement(is_functional)   ;{{{
{
  ; Select EA
  GroupAdd, ea_windows, ahk_class Afx:00400000:b:00010013:00000006:01A90B95
  GroupActivate, ea_windows, R

  ; Send keystroke to add a requirement
  SendInput ^m

  ; Select the type box and make it a requirement
  SendInput !y
  SendInput Req

  ; Select the stereotype box and make it the right value
  SendInput !r

  if(is_functional)
     SendInput Functional
  else
     SendInput Non-Func
  
  ; Clean up the clipboard and fill in the name box
  DirtyRequirementText = %clipboard%
  StringReplace, RequirementText, DirtyRequirementText, `r,,All
  StringReplace, RequirementText, DirtyRequirementText, `n,,All

  SendInput !m
  SendInput %RequirementText%

  ; Close the requirements box
  SendInput !o
} ;}}}




;**************************************************************************
; Timed Functions                                                      {{{
;**************************************************************************
; Automatically reload this script if changed
UPDATEDSCRIPT:
FileGetAttrib,attribs,%s_actual_script%
IfInString,attribs,A
{
   FileSetAttrib,-A,%s_actual_script%
   TrayTip, Rex's Macros, Updated Script!, 5, 1
   Sleep,2000
   Reload
}
Return 

; Automatically reload this script's configuration if changed
UPDATEDCONFIG: ;
FileGetAttrib,attribs,%config_file%
IfInString,attribs,A
{
   FileSetAttrib,-A,%config_file%
   TrayTip, Rex's Macros, Reloading Config!, 5, 1
   Sleep,2000
   Reload
}
Return  ;

PREVENTSCREENSAVER:
IfGreater, A_TimeIdle, 90000
{
   Send {Shift Down}{Shift Up}
   TrayTip, IDLE, Sent Keystroke, 1, 1
   Sleep 300
   TrayTip
}
Return

KillAutoUpdatesDlg:
SetTitleMatchMode,3
IfWinExist, Automatic Updates
{
   WinActivate
   ControlClick, Button2
}
Return

;                                                                       }}}
;**************************************************************************
; vim:nowrap:fdm=marker:tw=0
