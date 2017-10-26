﻿#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir   %A_ScriptDir% ; Ensures a consistent starting directory.

/*
New from 1.1
Changed code a little bit to make some functions. 
Added a 1-Up Feature
reload Macro with saving effectivity
Added a format change to handle semicolons
*/

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Autorun section \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


#KeyHistory 0
SetBatchLines, -1
SetDefaultMouseSpeed, 0
SetWinDelay, 0
SetControlDelay, 0
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen
#SingleInstance Force
++SetTitleMatchMode, 2
DetectHiddenWindows On
DetectHiddenText on
#include <Gdip>
#include <Gdip_ImageSearch>
#InstallKeybdHook
#InstallMouseHook
;OnExit, Exit_Label
ListLines Off
Version_Number = 1.2

Global Serialz, Clippy, Title, sleepstill, currMon

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Startup information \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

#include serials_startup.ahk

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Menu Tray \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Menu, Tray, NoStandard 
Menu, Tray, Add, How to use, HowTo
Menu Tray, Add, Check For update, Versioncheck
Menu, Tray, Add, Quit, Quitapp


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. \./.\./. Run update checker and start Gui screen \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

IfExist C:\SerialMacro\TempAdd.txt 
{
FileRead, editfield, C:\SerialMacro\TempAdd.txt 
FileDelete, C:\SerialMacro\TempAdd.txt 
}

IfExist C:\SerialMacro\TempAdded.txt 
{
FileRead, editfield2, C:\SerialMacro\TempAdded.txt
FileDelete,C:\SerialMacro\TempAdded.txt 
}

IfExist  C:\SerialMacro\Tempamount.txt
{
FileRead, totalprefixes,  C:\SerialMacro\Tempamount.txt
FileDelete, C:\SerialMacro\Tempamount.txt
}

gosub, Updatechecker
gosub, SerialsGUIscreen


pToken := Gdip_Startup()
 bmpNeedle1 := Gdip_CreateBitmapFromFile(FileNamered)
 bmpNeedleSearch := Gdip_CreateBitmapFromFile(FileNameSearch)
 bmpNeedleCheck := Gdip_CreateBitmapFromFile(FileNameCheck)
 bmpNeedleB := Gdip_CreateBitmapFromFile(FileNameButton)
/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Ctrl 2 start Macro \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

$^2::
{
GuiControlGet, editfieldcheck,,EditField
if editfieldcheck = 
{
	activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "No Prefixes added"
Settimer, winmovemsgbox, 20
msgbox,262144, Danger Will Robinson!!! -- Effectivity Macro V%Version_Number%, No Prefixes added
Exit
}		
Gosub, Enterallserials
skipbox = 0
}

^q::
{
Gosub, Exitprogram
Return
}
/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Enter Serials \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Enterallserials:
{
IniRead, refreshrate,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate																		  
;Refreshrate := 20
Prefixcount = 5
addtime = 0
;Badlist = Temp 
Badlist = 
StartTime := A_TickCount

activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Click on the ACM window that you want to add effectivity to and then press the OK button"
Settimer, winmovemsgbox, 20
Msgbox,262144,, Click on the ACM window that you want to add effectivity to and then press the OK button
Sleep 500
WinGetTitle, Title, A
Monitorprogram := Title
sleep 300
ACMscreen := GetCurrentMonitor()

GOsub, SerialFullScreen

						   
						   
						   
						   

						   
;msgbox,262144,,Please maximize the ACM window to help prevent problems.`n`n Press the OK button once that is complete.
Prefixcount = 5
Badlist = 

Guicontrol,hide, Starting

Gui, Submit, NoHide

Stoptimer = 0
 Serialzcounter2 = 0
Serialzcounter = 0


Gosub, Getmousepositions

;Msgbox, MOnitor is %currmon%

Gosub, Comma_Check 
Gosub, Searchend

	Loop
	{
IniRead, refreshrate,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate																		   
	Gosub, checkforactivity ; check every second (100ms) how long there has been no activity
	
		;**********************************
		;**** For when Esc is pressed *****
		;**********************************
		
		If breakloop = 1
		{
		Gui 1: -AlwaysOnTop
		Break
		SplashTextOn,,,Macro Stopped		
		Guicontrol,hide, Start
		Guicontrol,hide, paused
		Guicontrol,show, Stopped
		Guicontrol,hide, Running
		Gui, Submit, NoHide		
		}

		Gosub, loopcounts
		If LoopCount >= %Refreshrate%
			;If Needrefresh = 1
			{
				Click, %Applyx%,%Applyy%
				Click, %Applyx%,%Applyy%
				Sleep 1000
				;SlowACM = %searchcount%
				Needrefresh = 0
				;msgbox, refresh
				sleep 300
				Gosub, Searchend
				Gosub, Win_check
				sleep 100
				Send {F5}
				Sleep 2000
				Searchcount = 0
				Searchcountser = 0
				Gosub Refreshpage
				Loopcount = 0
				Searchcount = 0
				Searchcountser = 0
				Needrefresh = 0
				Sleep 1000
			}
		
		
	Sleep 100
	Serialz=0
	Searchcount = 0
	Searchcountser = 0
	;Settimer, checkforactivity, Off
	Gosub, Get_Prefix
	Gosub, Copy_Serial
	Gosub, Number_Check
	}
return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Get mouse positions ./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Getmousepositions:
{

SetTimer, ToolTipTimerbutton,Off
settimer, ToolTipTimerprefix,Off
SetTimer,ToolTipTimerapply,off
tooltip,

SetTimer, ToolTipTimerbutton, 10  ;timer routine will occur every 10ms..

Keywait, Shift,D
Keywait, Lbutton, D							  

MouseGetPos, Searchx, Searchy


WinGetTitle, Title, A
;WinSet, AlwaysOnTop,on,%Title%
currMon := GetCurrentMonitor()
Sleep 500
tooltip,
SetTimer,ToolTipTimerbutton,off

Sleep 1000

settimer, ToolTipTimerprefix,10


Keywait, Shift,D
Keywait, Lbutton, D							  

MouseGetPos, prefixx, prefixy
tooltip,


settimer, ToolTipTimerprefix,Off
Sleep 1000


SetTimer, ToolTipTimerapply, 10  ;timer routine will occur every 10ms..


Keywait, Shift,D
Keywait, Lbutton, D						  


MouseGetPos, Applyx, Applyy
tooltip,
SetTimer,ToolTipTimerapply,off


Sleep 500
ToolTip, 
Click, %Searchx%, %Searchy%
Sleep 500
Click, %prefixx%, %prefixy%
Sleep 1000
mousemove, 300,300

Return
}




/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Get Prefix \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Get_Prefix:
{
;Guicontrol, Focus, Editfield
Modifier = 
Guicount++
COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, Effectivity Macro 
sleep 50
COntrolSEnd,,{Shift Down}{Right 3}{Shift Up}, Effectivity Macro 

ControlGet,Prefixes,Selected,,,Effectivity Macro 

sleep 300

If Prefixes = ***
{
complete = 1
}


COntrolSEnd,, {del}, Effectivity Macro 

Text1 = %Prefixes%
;Msgbox, Prefix is %text1% 

;Guitext%Guicount% = %Text1%
Guitextlocation = %Prefixes%

if Prefixcount != 5
{
PrefixStore1 = %PrefixStore%
Prefixstore =  %Guitextlocation%
;msgbox, prefix store is %prefixstore% and prefixstore1 is %PrefixStore1%
}
If Prefixcount = 5
{
PrefixStore1 = %Guitextlocation%
Prefixstore = %Guitextlocation%
Prefixcount = 1
;msgbox, prefix store is %prefixstore%
}

Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Copy Serial  \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Copy_Serial:
{

	
Serialz++
COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, Effectivity Macro
sleep 50
COntrolSEnd,,{Shift Down}{Right 5}{Shift Up},Effectivity Macro 

ControlGet, SerialN, Selected,,,Effectivity Macro 

COntrolSEnd,, {del}, Effectivity Macro 
	If serialz = 1
	{
		Clippy = %SerialN%


		If Serialzcounter  = 0
		{
		If Serialstore1 = 
		{
		Serialstore1 = %Clippy%
		}
		Serialstore = %Clippy%
		Serialzcounter = 1
		}
		Else 
		{
		Serialstore1 = %Serialstore%
		Serialstore =  %Clippy%
		}
	}
	
	If Serialz = 2
	{
	Clippy2 = %SerialN%

		If Serialzcounter2  = 0
		{
		If Serialstore3 = 
		{
		Serialstore3 = %Clippy2%
		}
		Serialstore2 = %Clippy2%
		Serialzcounter2 = 1
		}
		Else 
		{
		Serialstore3 = %Serialstore2%
		Serialstore2 =  %Clippy2%
		}
		COntrolSEnd,, {del 2}, Effectivity Macro 
		GOsub, Upcheck
	}

Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Loop counts \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

loopcounts:
{
LoopCount++
return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Number checks \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Number_Check:
{

;2rd set of numbers check ===========
COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, Effectivity Macro
sleep 50
COntrolsend,,{Shift down}{RIght}{SHift up},Effectivity Macro  

ControlGet,Numbers,Selected,,,Effectivity Macro 

COntrolSEnd,, {del},Effectivity Macro 

Text3 = %Numbers%

;msgbox, text is %text3%


If Text3 = -
{ 


Gosub, Copy_Serial
If Clippy2 =
	{	
		Clippytemp = 99999
	}
	else {
	clippytemp = %clippy2%
	}
	
Searchcount = 0
Searchcountser = 0
Addser = %Prefixes%%Clippy%
nextserialtoaddv = %Addser%-%Clippytemp%	
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
Gui, submit, nohide
Gosub, Searchendforserials	
GOsub, EnterSerials
;gosub, Number_Check
Return
}

Else If Text3 = ,
{
		
	If Serialz = 1
	{
	Clippy2 = %Clippy%
	Searchcount = 0
	Searchcountser = 0
	Addser = %Prefixes%%Clippy%
	nextserialtoaddv = %Addser%-%Clippy2%
	GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
	COntrolSEnd,, {del},Effectivity Macro 
	Gui, submit, nohide
	Gosub, Searchendforserials
	GOsub, EnterSerials
	Return
	}
Else if Serialz = 2
{
Addser = %Prefixes%%Clippy%
nextserialtoaddv = %Addser%-%Clippy2%
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
COntrolSEnd,, {del},Effectivity Macro 
Gui, submit, nohide
Gosub, Searchendforserials
GOsub, EnterSerials
Return
}

}

Else if Text3 = 
{

If Serialz = 1
{

	Clippy2 = %Clippy%
	
	If Clippy2 =
	{	
		Clippytemp = 99999
	}
	else {
	clippytemp = %clippy2%
	}	


	Searchcount = 0
	 Searchcountser = 0
	Addser = %Prefixes%%Clippy%
	nextserialtoaddv = %Addser%-%Clippytemp%	

if complete = 1
GuiControl,1:,nextserialtoadd, 
Else
	GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
	COntrolSEnd,, {del},Effectivity Macro 
	Gui, submit, nohide
	
Gosub, Searchendforserials
GOsub, EnterSerials
Return
}	
	Addser = %Prefixes%%Clippy%
	nextserialtoaddv = %Addser%-%Clippytemp%	
GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
COntrolSEnd,, {del},Effectivity Macro 
Gui, submit, nohide


Gosub, Searchendforserials
Gosub, EnterSerials	
Return
}

return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Comma check \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/ 

Comma_Check:
{

COntrolSEnd,, {Shift Down}{Right}{SHift Up},Effectivity Macro 
ControlGet,Commacheck,Selected,,,Effectivity Macro 
If Commacheck = ,
{
;Msgbox, COmmafound
Controlsend,,{BackSpace 2},Effectivity Macro 
Return
}

Else 
{
;msgbox, comma not found
COntrolSEnd,, {Left}{BackSpace 2}, Effectivity Macro 
Return
}
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Check space \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Check_Space:
{
If Clippy = 
{
ControlSEnd,,{Del},Effectivity Macro 
ExitApp
}
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Reloading  \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Reloading:
{
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Press Okay Button to reload"
Settimer, winmovemsgbox, 20						   
Send {Shift Up}{Ctrl Up}
Msgbox,262144,Reload --Effectivity Macro V%Version_Number%, The number of successful Serial additions to ACM is %Serialcount% `n`n Serial Macro is currently Paused.`n`n Press Okay Button to reload Macro.`n`n Be sure to View the Serials Already Added Section to view the  Serial Numbers that were added into ACM to figure out where to start over if the macro messed up.
Pause, on
Reload
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. ESC \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/ 

#IfWinActive ahk_class TTAFrameXClass
Esc::
{

Gosub,Exitprogram 
Return
}

#ifwinactive Effectivity Macro V
Esc::
{

Gosub,Exitprogram 
Return
}

#ifwinactive

Exitprogram:
{

activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Are you sure that you want to stop the macro"
Settimer, winmovemsgbox, 20					   
Msgbox,262148,Stop -- Effectivity Macro V%Version_Number%, The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure that you want to stop the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.
ifmsgbox yes
{
Stopactcheck = 1						
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Send {Shift Up}{Ctrl Up}
breakloop = 1
Exit
}
Else 
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Up check \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

UPcheck:
{
If Clippy2 = 99999
Clippy2 = %A_Space%

Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Quit app  \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/
Quitapp:
{
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Are you sure you want to quit"
Settimer, winmovemsgbox, 20
msgbox,262148,Quit -- Effectivity Macro V%Version_Number%, Are you sure you want to quit?
ifMsgBox Yes
{
Stopactcheck = 1					
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Send {Shift Up}{Ctrl Up}
breakloop = 1
WinSet, AlwaysOnTop, Off,%Title%
ExitApp
}
Else
{
Return
}
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Win check \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Win_check:
{
IfWinNotActive , %Title%
{
WinActivate, %Title%
WinWaitActive, %Title%,,3
Sleep 500
}

return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Search end \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Searchend:
{
;msgbox, searchend
listlines off
bmpHaystack := Gdip_BitmapFromScreen(currMon)
Sleep 100
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
sleep 100
;listlines on
If RETSearch < 0
{
if RETSearch = -1001
RETSearch = invalid haystack or needle bitmap pointer
if RETSearch = -1002
RETSearch = invalid variation value
if RETSearch = -1003
RETSearch = Unable to lock haystack bitmap bits
if RETSearch = -1004
RETSearch = Unable to lock needle bitmap bits
if RETSearch = -1005
RETSearch = Cannot find monitor for screen capture
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Error Searchend (bmpNeedle1)"
Settimer, winmovemsgbox, 20
Msgbox,262144, Effectivity Macro V%Version_Number%, Error Searchend (bmpNeedle1) %RetSearch%
Exit
}

If RETSearch > 0
{
SetTimer, refreshcheck, Off
Refreshchecks = 0
SleepStill = 0 
;Msgbox, found
}



If RETSearch = 0
{
If Searchcount = 7
{
Timeout = Yes
gosub, Macrotimedout
}
;Msgbox, notfound
Sleepstill = 1
Sleep 500
searchcount++
Gosub, Searchend
}
Return
}


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Searchend for serials \./.\./.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/ 

Searchendforserials:
{
/*
msgbox,4,,move forward is yes`n go to timeout is no
ifmsgbox no
gosub, Macrotimedout
*/
/*
msgbox,4,,yes for found, no for not found
ifmsgbox yes
RETSearch = 1
ifmsgbox no
RETSearch = 0
Searchcountser = 5

*/
 		
;msgbox, end serials
listlines off
bmpHaystack := Gdip_BitmapFromScreen(currMon)
Sleep 100
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
sleep 100
;listlines on
If RETSearch < 0
{
if RETSearch = -1001
RETSearch = invalid haystack or needle bitmap pointer
if RETSearch = -1002
RETSearch = invalid variation value
if RETSearch = -1003
RETSearch = Unable to lock haystack bitmap bits
if RETSearch = -1004
RETSearch = Unable to lock needle bitmap bits
if RETSearch = -1005
RETSearch = Cannot find monitor for screen capture
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Error Searchendserials (bmpNeedle1)" RetSearch
Settimer, winmovemsgbox, 20
Msgbox,262144, Effectivity Macro V%Version_Number%, Error Searchendserials (bmpNeedle1) %RetSearch%
Exit
}


If RETSearch > 0
{
SplashTextOff
Return
}

If RETSearch = 0
{
;msgbox, notfound serialsearch

listlines off

If addtime = 2
{

If Searchcountser = 1
{
SplashTextOn,,25,Serial Macro, Time out in 7
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 2
{
SplashTextOn,,25,Serial Macro, Time out in 6
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}

If Searchcountser = 3
{
SplashTextOn,,25,Serial Macro, Time out in 5
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}

If Searchcountser = 4
{
SplashTextOn,,25,Serial Macro, Time out in 4
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}


If Searchcountser = 5
{
SplashTextOn,,25,Serial Macro, Time out in 3
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 6
{
SplashTextOn,,25,Serial Macro, Time out in 2
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
If Searchcountser = 7
{
SplashTextOn,,25,Serial Macro, Time out in 1
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
}
;listlines on
If Searchcountser = 8
{
SplashTextoff
Timeout = Yes
gosub, Macrotimedout
}
}


If addtime = 1
{

	If Searchcountser = 1
	{
	SplashTextOn,,25,Serial Macro, Time out in 6
	Click, %Applyx%,%Applyy%
	Click, %Applyx%,%Applyy%
	}

	If Searchcountser = 2
	{
	SplashTextOn,,25,Serial Macro, Time out in 5
	Click, %Applyx%,%Applyy%
	Click, %Applyx%,%Applyy%
	}

	If Searchcountser = 3
	{
	SplashTextOn,,25,Serial Macro, Time out in 4
	Click, %Applyx%,%Applyy%
	Click, %Applyx%,%Applyy%
	}


	If Searchcountser = 4
	{
	SplashTextOn,,25,Serial Macro, Time out in 3
	Click, %Applyx%,%Applyy%
	Click, %Applyx%,%Applyy%
	}
	If Searchcountser = 5
	{
	SplashTextOn,,25,Serial Macro, Time out in 2
	Click, %Applyx%,%Applyy%
	Click, %Applyx%,%Applyy%
	}
	If Searchcountser = 6
	{
	SplashTextOn,,25,Serial Macro, Time out in 1
	Click, %Applyx%,%Applyy%
	Click, %Applyx%,%Applyy%
	}
	;listlines on
	If Searchcountser = 7
	{
	SplashTextoff
	Timeout = Yes
	gosub, Macrotimedout
	}
}
	Else 
	{
		If Searchcountser = 1
		{
		SplashTextOn,,25,Serial Macro, Time out in 5
		Click, %Applyx%,%Applyy%
		Click, %Applyx%,%Applyy%
		}

		If Searchcountser = 2
		{
		SplashTextOn,,25,Serial Macro, Time out in 4
		Click, %Applyx%,%Applyy%
		Click, %Applyx%,%Applyy%
		}


		If Searchcountser = 3
		{
		SplashTextOn,,25,Serial Macro, Time out in 3
		Click, %Applyx%,%Applyy%
		Click, %Applyx%,%Applyy%
		}
		If Searchcountser = 4
		{
		SplashTextOn,,25,Serial Macro, Time out in 2
		Click, %Applyx%,%Applyy%
		Click, %Applyx%,%Applyy%
		}
		If Searchcountser = 5
		{
		SplashTextOn,,25,Serial Macro, Time out in 1
		Click, %Applyx%,%Applyy%
		Click, %Applyx%,%Applyy%
		}
		;listlines on
		If Searchcountser = 6
		{
		SplashTextoff
		Timeout = Yes
		gosub, Macrotimedout
		}
	}

;Msgbox, notfound
Sleepstill = 1
Sleep 500
searchcountser++
Gosub, Searchendforserials
}
Return
}






/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Refresh check \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/ 
refreshcheck:
{
Refreshchecks++
 If refreshchecks = 10
 {
 Needrefresh = 1
 Refreshchecks = 0
 }
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Get current monitor \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

GetCurrentMonitor()
{
SysGet, numberOfMonitors, MonitorCount
WinGetPos, winX, winY, winWidth, winHeight, A
winMidX := winX + winWidth / 2
winMidY := winY + winHeight / 2
Loop %numberOfMonitors%
{
SysGet, monArea, Monitor, %A_Index%
if (winMidX > monAreaLeft && winMidX < monAreaRight && winMidY < monAreaBottom && winMidY > monAreaTop)
return A_Index
}
SysGet, MonitorPrimary, MonitorPrimary
return "No Monitor Found"
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Refresh page \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/ 

Refreshpage:
{
Sleep 500
Click, %Searchx%, %Searchy%
Sleep 2000
Gosub, SearchendRefresh
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Search end refresh \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

SearchendRefresh:
{
bmpHaystack := Gdip_BitmapFromScreen(currMon)
Sleep 100
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
sleep 100

If RETSearch < 0
{
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Error Searchend (bmpNeedle1)"
Settimer, winmovemsgbox, 20
Msgbox,262144,Effectivity Macro V%Version_Number%, Error SearchendRefresh (bmpNeedle1) %RetSearch%
Exit
}

If RETSearch > 0
{
SleepStill = 0 

Return
;Msgbox, found
}

If RETSearch = 0
{
;Msgbox, notfound
searchcount++
Gosub, Refreshpage
}
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Screen check \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/ 
Screen_check:
{
listlines off
bmpHaystack1 := Gdip_BitmapFromScreen(currMon)
ScreenSearch := Gdip_ImageSearch(bmpHaystack1,bmpNeedleSearch,,0,0,0,0,5,0,0,0)

If ScreenSearch < 0
{
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Error screensearch (bmpNeedleSearch)"
Settimer, winmovemsgbox, 20
Msgbox,262144, Effectivity Macro V%Version_Number%, Error screensearch (bmpNeedleSearch) is  %ScreenSearch%
Reload
}

If ScreenSearch > 0
{
Return
}

If ScreenSearch = 0  
{
ScreenCheck := Gdip_ImageSearch(bmpHaystack1,bmpNeedleCheck,,0,0,0,0,5,0,0,0)
If ScreenCheck < 0
{
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Error screensearch (bmpNeedleCheck)"
Settimer, winmovemsgbox, 20
Msgbox,262144, Effectivity Macro V%Version_Number%, Error  screen check  (bmpNeedleCheck) is %ScreenCheck%
Reload
}
If ScreenCheck > 0
{
Return
}

;listlines on

If ScreenCheck = 0
{
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Make the effectivity screen active and what is most likely something over the Effectivity screen"
Settimer, winmovemsgbox, 20
Msgbox,262144,Effectivity Macro V%Version_Number%, OOPS!!`n`n Something bad happened and Cannot find the Effectivity screen.  `n`nMacro is now paused. `n`nMake the effectivity screen active and what is most likely something over the Effectivity screen. Hide that panel and find where I left off. `n`n Press Pause to unpause script after you press OK on this window.
Pause, on
}
Return
}

Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Enter serials  \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Enterserials:
{

	
Gosub, Win_check


		
Loop, parse, Badlist, `,,all
{
;msgbox, loopfield is %A_LoopField%
If prefixes = %A_LoopField%
{
;msgbox, bad acm  match in enter serials to %A_LoopField%
noacmPrefix = %prefixes%
Skipserial = 1
Break
}
else
Skipserial = 0
}

Loop, parse, DualENG, `,,all
{
;msgbox, loopfield is %A_LoopField%

If prefixes = %A_LoopField%
{
;msgbox,  serial number match
DUalACMCheck = 1
DualACMPrefix = %prefixes%
Break
}
Else 
{
;msgbox, no match
DUalACMCheck = 0
}
}


if Skipserial = 1
{	
	;msgbox, no acm skip
	Modifier = --- *Serial Not in ACM*
	Guicontrol,1:,Editfield2, %Editfield2%%PrefixStore%%Clippy%-%Clippy2%%Modifier%`n
	Guicontrol,hide, Editfield2,
	GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%	
	Modifier = 
	Skipserial = 0
	;COntrolSEnd,, {del 2},Enter Serial Macro 
	return
}	

If Complete = 1
		{
		ElapsedTime := A_TickCount - StartTime
		res1 := milli2hms(ElapsedTime, h, m, s)
		Sleep 500
		Send {f5}
		activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Macro Finished due to no more Serials to add"
Settimer, winmovemsgbox, 20
		Msgbox,262144,Effectivity Macro V%Version_Number%, The number of successful Serial additions to ACM is %Serialcount% `n`n Macro Finished due to no more Serials to add. `n`n It took the macro %res1% to perform tasks. `n`n Please close Serial Macro Window when finished checking to ensure serials were entered correctly.
		Guicontrol,1:, Editfield,
		gosub, radio2h
		Guicontrol,hide, Start
		Guicontrol,hide, paused
		Guicontrol,show, Stopped
		Guicontrol,hide, Running
		Exit
		Return
		}
		
;listlines on
Gosub, Screen_Check
Click, %prefixx%, %prefixy%
sleep 100
mousemove 300,300
sleep 100
SEndRaw, %text1%
Sleep 100
Send {Tab}
Gosub, Win_check
Sendraw, %Clippy%
sleep 250
Send {Tab}
Gosub, Win_check
SendRaw, %Clippy2%
Sleep 250
Send {Tab}
Sleep 650
	If DUalACMCheck = 1
	{
		Modifier = --- **Multiple Engineering Models** 
				activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Please select the model in ACM and press the OK button"
Settimer, winmovemsgbox, 20
		Msgbox,262144,-- Effectivity Macro V%Version_Number%, This Serial has multiple Engineering Models. Please select the model in ACM and press the OK button in this window.
		;Guicontrol,1:,Editfield2, %Editfield2%%Guitextlocation%%Clippy%-%Clippy2%%Modifier%`n
			SplashTextOn,,25,Serial Macro, Macro will resume in 2
			sleep 1000
			SplashTextOn,,25,Serial Macro, Macro will resume in 1
			sleep 1000
		SplashTextOff
		gosub, Win_check
		Click, %prefixx%, %prefixy%
		sleep 500
		Send {Tab 3}
		sleep 500
		DualACMCheck = 0
		SleepStill = 0 
		sleep 250
		
	}
	Send {enter 2}
	Click, %Applyx%,%Applyy%
	Click, %Applyx%,%Applyy%
	;SetTimer, refreshcheck, 250
	Serialcount +=1	
	Sleep 100
	Searchcount = 0
	Searchcountser = 0
	
	
	If Clippy2 =
	{	
		Clippy2 = 99999
	}
	Guicontrol,1:,Editfield2, %Editfield2%%PrefixStore%%Clippy%-%Clippy2%%Modifier%`n
	Guicontrol,hide, Editfield2,
	GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%
;SetTimer, refreshcheck, Off
Refreshchecks = 0
SleepStill = 0 
Modifier = 
Skipserial = 0
Return
}


milli2hms(milli, ByRef hours=0, ByRef mins=0, ByRef secs=0, secPercision=0)
{
  SetFormat, FLOAT, 0.%secPercision%
  milli /= 1000.0
  secs := mod(milli, 60)
  secs = %secs%Sec
  SetFormat, FLOAT, 0.0
  milli //= 60   
  mins := mod(milli, 60)
  mins = %mins%Min
  hours := milli //60 
  Hours = %hours%hrs	
  return hours . ":" . mins . ":" . secs
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Exit label \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/
EXIT_LABEL:
{
If pToken !=0
{
Gdip_Shutdown(pToken)
}
ExitApp
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.throwout  \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

throwout:
{
badserial = 1
Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.Tool tip timers \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

ToolTipTimerbutton:
{
tooltip, %Textaddbutton%
Return
}

ToolTipTimerapply:
{
tooltip, %Textapplybutton%
Return
}

ToolTipTimerprefix:
{
tooltip, %Textprefixbutton%
Return
}

KeystateTimer:
{
GetKeyState,Rmousee, Rbutton
GetKeyState,Lmousee, Lbutton

If Lmousee = D
{
Return 
}
 IF Rmousee = D
{
Return
}
Else
{
Sleep 100
GOsub, KeystateTimer
}
Return
}


Engmodel:
{
warned = 1
WinGetPos,ex,ey,,,Macro Timed Out
IniWrite, %ex%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
IniWrite, %ey%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition																	
Gui, 3:Destroy
Modifier = --- **Multiple Engineering Models** 
;DualACMPrefix = %PrefixStore1%
;msgbox, noprefix is %noacmPrefix%
DUalACMCheck = 1
DualENG = %DualENG%%PrefixStore1%`,
StringTrimRight, EditField2, Editfield2,1
;msgbox, after trim :`n %EditField2%`n Prefixstore is %PrefixStore1%`nSerial is %Serialstore1%-%Serialstore3% Mod is %Modifier%
Editfield2n = %Editfield2%%Modifier%`n
EditField2 = %EditField2n%
Guicontrol, 1:,Editfield2, %Editfield2%
Gui 1: +alwaysontop
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Select the engeering model and then press the OK button on this window. The prefix will be logged for this session to prevent more timouts of the same prefix."
Settimer, winmovemsgbox, 20					   
Msgbox,262144, Effectivity Macro V%Version_Number%, Select the engeering model and then press the OK button on this window. `n`n The prefix will be logged for this session to prevent more timouts of the same prefix.
sleep 100
gosub, Win_check
Click, %prefixx%, %prefixy%
Send {Tab 3}
SplashTextOn,,25,Serial Macro, Macro will resume in 3
sleep 250 
SplashTextOn,,25,Serial Macro, Macro will resume in 2 
sleep 250
SplashTextOn,,25,Serial Macro, Macro will resume in 1
sleep 250 

SplashTextOff
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, Off

Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
Click, %Applyx%,%Applyy%
modifier = 
gosub, searchend
Return
}

Anarchy:
{
WinGetPos,ex,ey,,,Macro Timed Out
IniWrite, %ex%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
IniWrite, %ey%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition
Gui, 3:Destroy
				activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := " No button to reload macro"
Settimer, winmovemsgbox, 20
Msgbox,262148,-- Effectivity Macro V%Version_Number%, Press Yes button if you found the issue and to Resume the Macro, No button to reload macro
IfMsgBox Yes
{
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, Off
}
Else
{
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,show, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
Pause, Off 
Gosub, Reloading
}
return
}

Serialnogo:
{
WinGetPos,ex,ey,,,Macro Timed Out
IniWrite, %ex%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
IniWrite, %ey%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition
Gui, 3:Destroy
Modifier := --- *Serial Not in ACM*
Serialcount--
;noacmPrefix = %PrefixStore1%
;msgbox, noprefix is %noacmPrefix%
Badlist = %badlist%%PrefixStore1%`,

;msgbox, badlist after nogo is %badlist%

SplashTextOn,,25,Serial Macro, Macro will resume in 3
sleep 500 
SplashTextOn,,25,Serial Macro, Macro will resume in 2 
sleep 500
SplashTextOn,,25,Serial Macro, Macro will resume in 1
sleep 500 
SplashTextOff


Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, 1:Submit, NoHide
Pause, Off
gosub, Win_check
Click, %prefixx%, %prefixy%
sleep 500
Send {ctrl down}{a}{Ctrl up}
sleep 100
Send {Del}{Tab}
Send {Del}{Tab}
Send {Del}{Tab}
Click, %prefixx%, %prefixy%
Send {Del}

Modifier = --- *Serial Not in ACM*
;msgbox, Beforetrim:`n %EditField2%
StringTrimRight, EditField2, Editfield2,1
;msgbox, after trim :`n %EditField2%`n Prefixstore is %PrefixStore1%`nSerial is %Serialstore1%-%Serialstore3% Mod is %Modifier%

Editfield2n = %Editfield2%%Modifier%`n
EditField2 = %EditField2n%


Guicontrol, 1:,Editfield2, %Editfield2%
Gui 1: +alwaysontop
Modifier =
Return
}

pausedstate:
{
Sleep 2000
return
}

checkforactivity:
{
 if A_TimeIdlePhysical < 4999 ; meaning there has been user activity
{
Gui 1: -AlwaysOnTop

{
	If Stopactcheck = 1
	{
	SplashTextOff
	Return
	}
	If (A_TimeIdlePhysical > 0) and (A_TimeIdlePhysical < 1000)
	Timeleft = 5
	
	If Stopactcheck = 1
	{
	SplashTextOff
	Return
	}
	If (A_TimeIdlePhysical > 1000) and (A_TimeIdlePhysical < 2000)
	Timeleft = 4
	
If Stopactcheck = 1
	{
	SplashTextOff
	Return
	}
	If (A_TimeIdlePhysical > 2000) and (A_TimeIdlePhysical < 3000)
	Timeleft = 3
	
	If Stopactcheck = 1
	{
	SplashTextOff
	Return
	}
	If (A_TimeIdlePhysical > 3000) and (A_TimeIdlePhysical < 4000)
	Timeleft = 2
	
	If Stopactcheck = 1
	{
	SplashTextOff
	Return
	}
	
	If (A_TimeIdlePhysical > 4000) and (A_TimeIdlePhysical < 5000)
	Timeleft = 1	
		
		If Stopactcheck = 1
	{
	SplashTextOff
	Return
	}}
SplashTextOn,350,50,Macro paused, Macro is now paused due to user activity.`n Macro will resume after %timeleft% seconds of no user input
pausedstate = yes
 Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
	Sleep 1000
	Gosub, checkforactivity
 } 
  if A_TimeIdlePhysical > 5000 ; meaning there has been no user activity
{
If pausedstate = yes
{
splashtextoff
Gui 1: +AlwaysOnTop
Gosub, Win_check
pausedstate = no
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
gosub, radio1h
Gui, Submit, NoHide
gosub, SerialFullScreen
sleep 1000
}
return
}
return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Pause\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Insert::
Pause::
{
if A_IsPaused = 0
{
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Press pause to unpause"
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,10
Pause, toggle, 1
Return
}
Else
{
gosub, radio1h
Gui 1: +AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, toggle, 1 
}
return
}

pausesub:
{
if A_IsPaused = 0
{
gosub, radio1h
Gui 1: -AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,show, paused
Guicontrol,hide, Stopped
Guicontrol,hide, Running
Gui, Submit, NoHide
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Press pause to unpause"
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,.1
Settimer, winmovemsgbox, 20
Msgbox,262144,-- Effectivity Macro V%Version_Number%, Macro is paused. Press pause to unpause,10
Pause, toggle, 1 
Return
}
Else
{
Gui 1: +AlwaysOnTop
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui, Submit, NoHide
Pause, toggle, 1 
}
return
}

SerialFullScreen:
{
WinGetPos, Xarbor,yarbor,warbor,harbor, %title%
CurrmonAM := GetCurrentMonitor()
SysGet,Aarea,MonitorWorkArea,%currmonAM%
WidthA := AareaRight- AareaLeft
HeightA := aareaBottom - aAreaTop
leftt := aAreaLeft - 4
topp := AAreaTop - 4
MouseGetPos mmx,mmy
If yarbor = %topp%
{
If xarbor = %leftt%
;Msgbox, win maxed
Return
}
Else
;msgbox, not maxed
CoordMode, mouse, Relative
MouseMove 300,10
Click
Click
Coordmode, mouse, screen
;MouseMove, mmx, mmy
return
}

acmlong:
{
pause, off
Gui 1: +AlwaysOnTop
gosub, radio1h
Guicontrol,hide, Start
Guicontrol,hide, paused
Guicontrol,hide, paused
Guicontrol,hide, Stopped
Guicontrol,show, Running
Gui 1:Submit, NoHide
WinGetPos,ex,ey,,,Macro Timed Out
IniWrite, %ex%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
IniWrite, %ey%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition																
Gui, 3:Destroy
Click %Applyx%, %Applyy%
Click %Applyx%, %Applyy%
if addtime = 0
	{
		Addtimetemp++
		If addtimetemp = 2
		{
		Addtime = 1
		 Addtimetemp = 0
		Return
		}
	}
If addtime = 1
{
Addtimetemp++
		If addtimetemp = 2
		{
		Addtime = 2
		Addtimetemp = 0
		Return
		}
}
If addtime = 2
{
Addtimetemp++
		If addtimetemp = 2
		{
		Refreshrate /= 2
		 Addtimetemp = 0
		 addtime = 3
		SplashTextOn,350,50,Refresh time reduced, Macro will refresh ACM every 10 Serials to help reduce the amount of timeouts from a slow ACM
		Sleep 5000
		SplashTextOff
		Return
		}
}
If addtime = 3
{
Return
}
Return
}

; FormatSerials.ahk below
#include FormatSerials.ahk

; Serials Gui.ahk below
#include Serials GUI.ahk

;Version check_serialmacro.ahk below
#include Version_check_serialmacro.ahk


activeMonitorInfo( ByRef aX, ByRef aY, ByRef aWidth,  ByRef  aHeight, ByRef mouseX, ByRef mouseY  )
{
CoordMode, Mouse, Screen
MouseGetPos, mouseX , mouseY
SysGet, monCount, MonitorCount
Loop %monCount%
{ 	SysGet, curMon, Monitor, %a_index%
if ( mouseX >= curMonLeft and mouseX <= curMonRight and mouseY >= curMonTop and mouseY <= curMonBottom )
{
aX      := curMonTop
ay      := curMonLeft
aHeight := curMonBottom - curMonTop
aWidth  := curMonRight  - curMonLeft
return
}
}
}


Pausescript()
{
Menu,Tray,Icon, % "C:\Serialmacro\icons\paused.ico", ,1
Pause,on
Return
}

UnPausescript()
{
Menu,Tray,Icon, % "C:\Serialmacro\icons\Serial.ico", ,1
Pause,off
Return
}

SetTitleMatchMode, 2
winmovemsgbox:
{
	SetTimer, WinMoveMsgBox, OFF 
	WinMove, %titletext% , Amonx, Amony 
	return
}
#`:: Listvars

+#`:: ListLines