#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir   %A_ScriptDir% ; Ensures a consistent starting directory.

/*
New from 1.2
Create and Export Effectivity to Excel sheet after formatting

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
#InstallKeybdHook
#InstallMouseHook
;OnExit, Exit_Label
;ListLines Off
Global Prefix_Number_Location_Check, First_Effectivity_Numbers, Title, sleepstill, Current_Monitor, Effectivity_Macro, Version_Number, Addser

Version_Number = 1.3 Beta

Effectivity_Macro :=  "Effectivity Macro V" Version_Number
;msgbox, %Effectivity_Macro%



/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Startup information \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/
; THis section sets some variables up for later use

Checkp=0
Sleepstill = 0
LoopCount = -1
ButtonLoop = 0
Serialcount = 0
Programend = 0
Needrefresh = 0
Editfield2 = 
Complete = 0
Modifier =
ToPause = 0
badserial = 0
Serialzcounter2 = 0
Serialzcounter = 0
Prefixcount = 5
TotalPrefixes = 0
Textaddbutton = Please Shift + mouse button click on the "Add Button" in the ACM effectivity screen to get its position.
Textprefixbutton = Please Shift + mouse button click in the "prefix" edit field in the ACM effectivity screen to get it's location.
Textapplybutton = Please Shift + mouse button click on the "Apply button" in the ACM effectivity screen to get it's location.
Radiobutton = 1



/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
								This part does the following:
								
1. Checks to see if certain files exist. If they do not exist,then it installs them to the correct spot 
2. Reads the INI file to get the User settings

\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
															
	CHeck the following link for info on the Folder creation : 	https://autohotkey.com/docs/commands/FileCreateDir.htm
	
	Check the following link for info on the INI read: https://autohotkey.com/docs/commands/IniRead.htm
	
	Check the following link for info on INI Write: https://autohotkey.com/docs/commands/IniWrite.htm
	
	Check the following link for info on Fileinstall: https://autohotkey.com/docs/commands/FileInstall.htm
								
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/



IfNotExist, C:\SerialMacro
{
   FileCreateDir, C:\SerialMacro
   sleep()
}

IfNotExist, C:\SerialMacro\icons
{
   FileCreateDir, C:\SerialMacro\icons
   sleep()
}
IfnotExist, C:\Serialmacro\Settings.ini
{
   FileAppend,  C:\Serialmacro\Settings.ini
   sleep(5)
   IniWrite, 20,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
   Sleep()
   IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
}

IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
IniRead, Sleep_Delay,  C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay

If Sleep_Delay = Error
{
IniWrite, 3,  C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay
IniRead, Sleep_Delay,  C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay
}

If Refreshrate = Error
{
   IniWrite, 20,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
   
   IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
}

FileInstall, C:\ArbortextMacrosn\serial macro\How to use Effectivity Macro.pdf, C:\SerialMacro\How to use Effectivity Macro.pdf,1

FileInstall, C:\ArbortextMacrosn\icons\serial.ico, C:\SerialMacro\icons\serial.ico,1
FileInstall, C:\ArbortextMacrosn\icons\paused.ico, C:\SerialMacro\icons\paused.ico,1

IfNotExist, C:\SerialMacro\redimage.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\redimage.png, C:\SerialMacro\redimage.png,1
}

IfNotExist, C:\SerialMacro\plussign.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\plussign.png, C:\SerialMacro\plussign.png,1
}

IfNotExist, C:\SerialMacro\Activeplus.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\activeplus.png, C:\SerialMacro\activeplus.png,1
}

IfNotExist, C:\SerialMacro\orange button.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\orange button.png, C:\SerialMacro\orange button.png,1
}

IfNotExist, C:\SerialMacro\paused.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\paused.png, C:\SerialMacro\paused.png,1
}

IfNotExist, C:\SerialMacro\start.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\start.png, C:\SerialMacro\start.png,1
}

IfNotExist, C:\SerialMacro\Running.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\Running.png, C:\SerialMacro\Running.png,1
}

IfNotExist, C:\SerialMacro\Stopped.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\Stopped.png, C:\SerialMacro\Stopped.png,1
}

IfNotExist, C:\SerialMacro\background.png
{
   FileInstall, C:\ArbortextMacrosn\serial macro\background.png, C:\SerialMacro\background.png,1
}

IfnotExist C:\SerialMacro\uredimage.png
{
   FileNamered := "C:\SerialMacro\redimage.png"
}else  {
   FileNamered := "C:\SerialMacro\uredimage.png"
}

IfnotExist C:\SerialMacro\uplussign.png
{
   FileNameSearch := "C:\SerialMacro\plussign.png"
}else  {
   FileNameSearch := "C:\SerialMacro\uplussign.png"
}

IfnotExist C:\SerialMacro\uActiveplus.png
{
   FileNameCheck := "C:\SerialMacro\Activeplus.png"
}else  {
   FileNameCheck := "C:\SerialMacro\uActiveplus.png"
}

IfNotExist C:\SerialMacro\uorange button.png
{
   FileNameButton := "C:\SerialMacro\orange button.png"
}else  {
   FileNameButton := "C:\SerialMacro\uorange button.png"
}              


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
									This section creates the Tray Icon Menu
									GO to following link for more info: https://autohotkey.com/docs/commands/Menu.htm
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Menu, Tray, NoStandard 
Menu, Tray, Add, How to use, HowTo
Menu Tray, Add, Check For update, Versioncheck
Menu, Tray, Add, Quit, Quitapp


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
											This section Does the following:
1. Checks for temp files in the case the macro was reloaded with Effectivity
2. Goes to the Update checker
3. Starts up the main Program window

GO to following link for more info on fileread: https://autohotkey.com/docs/commands/FileRead.htm
GO to following link for more info on filedelete: https://autohotkey.com/docs/commands/FileDelete.htm

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

gosub, Updatechecker ; goes to the subroutine to check for updates
gosub, SerialsGUIscreen ; Goes to the subroutine to show the main Program screen

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
											This section Does the following:
1. atarts the funciton for the image searching
2.Puts the images into memory to be used for later on for the image searches

FOr more info go to Gdip_CreateBitmapFromFile section on the bottom half of the script

\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/



return


 /*
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* Format Serials Section*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*/

  ;Sets the hotkey for Ctrl + 1 or Ctrl + numpad 1
$^Numpad1::
$^1::
{
   Gosub, FormatSerialsMacro ; Goes to the Formatserials subroutine
   return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
								Ctrl 2 start Macro
								
THis section does the following:
1. Checks to see if the Effectivity editbox is empty and then starts macro if it has text in it.

For more info on GUIcontrolget go to https://autohotkey.com/docs/commands/GuiControlGet.htm
FOr more info on Settimer go to https://autohotkey.com/docs/commands/SetTimer.htm
for more info on msgbox go to https://autohotkey.com/docs/commands/MsgBox.htm
								
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/




$^2::
{
   GuiControlGet, editfieldcheck,,EditField
   if editfieldcheck = 
   {
      Move_Message_Box("262144","Danger Will Robinson!!!","No Prefixes added")
      ;~ msgbox,262144, Danger Will Robinson!!! %Effectivity_Macro%, No Prefixes added
      Exit
   }		
   ;~ Gosub, Enterallserials
   Enterallserials()
   skipbox = 0
}


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
								Ctrl Q Quit Macro
								
THis section does the following:
1. Goes to the Exitprogram subroutine
								
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

^q::
{
   Gosub, Exitprogram
   Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Enter Serials 
	THis section does the following:
1. Reads the INI file for the user selected refresh rate
2. asks the user to click on the ACM screen so that it can grab the title to check for full screen 
3. IF not full screen, make acm full screen
4. Goes to the Getmousepositions subroutine
5. Goes to the comma_check subroutine
6. Goes to the Searchend subroutine
7. Starts a loop to enter the serial Numbers into ACM

	Check the following link for info on the INI read: https://autohotkey.com/docs/commands/IniRead.htm
	Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm
	Check the following link for info on the WinGetTitle: https://autohotkey.com/docs/commands/WinGetTitle.htm
	Check the following link for info on the GuiControl: https://autohotkey.com/docs/commands/GuiControl.htm
	Check the following link for info on the Gui: https://autohotkey.com/docs/commands/Gui.htm
	Check the following link for info on the splashtext: https://autohotkey.com/docs/commands/SplashTextOn.htm
	
	
	
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Enterallserials()
{
   global 
   ; Initial setup before the macro starts to output to the screen
   
   IniRead, refreshrate,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate																		  
 
   Prefixcount = 5
   addtime = 0  
   Badlist = 
    Stoptimer = 0
   Serialzcounter2 = 0
   Serialzcounter = 0
     StartTime := A_TickCount
   Move_Message_Box("262144","Click on the ACM window that you want to add effectivity to and then press the OK button","Select ACM Screen")
   ;~ Msgbox,262144,, Click on the ACM window that you want to add effectivity to and then press the OK button
   sleep(5)
   WinGet,  Active_ID, ID,A
   Monitorprogram := Title
   sleep(3)
   
SerialFullScreen(Active_ID)   
GUI_Image_Set ("Run") ; options are Stop, Run, Pause, Start   
  
   Gosub, Getmousepositions
   gosub, Createtab
    
   Comma_Check(Effectivity_Macro)

Result :=   Searchend()
          If (Result = Failure) or (Result = Timedout)
            Exit
   
    Tabcount = 0
   runcount = 21
   
   Loop
   {
      tabcount++
      IniRead, refreshrate,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate																		   
      checkforactivity()  
      
      	  ;~ If Runcount > 20
	  ;~ {
	  ;~ runcount = 1
	  ;~ Gosub, GetSerial
	  ;~ Gosub, Copy_SerialTester
	  ;~ }
	  
	  Send {ctrl down} %Tabcount% {Ctrl up}
	  
	  if tabcount = 2
	  tabcount = 0
      
      ;**********************************
      ;**** For when Esc is pressed *****
      ;**********************************
      
      If breakloop = 1
      {
         Gui 1: -AlwaysOnTop    
         SplashTextOn,,,Macro Stopped	
         GUI_Image_Set ("Stop") ; options are Stop, Run, Pause, Start      
         Sleep(10)
          SplashTextOff
         Break
      }
      
      LoopCount++
      If LoopCount >= %Refreshrate%
       {
         Click, %Applyx%,%Applyy%
         Click, %Applyx%,%Applyy%
         sleep(10)
          Needrefresh = 0
         ;msgbox, refresh
         sleep(3)
         Result :=   Searchend()
          If (Result = Failure) or (Result = Timedout)
            Exit
            
        Win_check(Active_ID)
         sleep()
         Send {F5}
         sleep(20)
             Searchcountser = 0
         Gosub Refreshpage
         Loopcount = 0
         Searchcountser = 0
         Needrefresh = 0
         sleep(10)
      }
      
      
      sleep()
      Prefix_Number_Location_Check=0
       Searchcountser = 0
      ;Settimer, checkforactivity, Off
      Modifier = 
     Prefiix := Get_Prefix()
     If Prefix = Complete
   {
      Complete = 1      
      Enterserials()
   }
   
   First_Effectivity_Numbers :=  Copy_Serial("1") ; options are 1  for first set of Prefix numbers, 2 for second set of prefix numbers
      Result := Number_Check()
      
      If Result = -
         Second_Effectivity_Numbers := Copy_Serial("2") ; options are 1  for first set of Prefix numbers, 2 for second set of prefix numbers
   else if (Result = ,) or (Result = )
         {
          Second_Effectivity_Numbers = %First_Effectivity_Numbers%    
         Searchcountser = 0
          nextserialtoaddv = %Addser%-%Second_Effectivity_Numbers%
         GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
         COntrolSEnd,, {del},%Effectivity_Macro% 
         Gui, submit, nohide
         }
 
Result :=   Searchend()
          If (Result = Failure) or (Result = Timedout)
            Exit
            
      GOsub, EnterSerials
  
   }
   return
}

createtab:
{
sleep(3)
Send {Alt Down}{d}
sleep()
Send {Enter}{Alt Up}
Sleep(20)

Click, %Effx%,%Effy%
Sleep(10)
click, %Add_Button_X_Location%, %Add_Button_Y_Location%
Sleep(10)
Result := SearchendRefresh()
 If (Result = Failure) or (Result = Timedout)
            Exit   
Click, %prefixx%, %prefixy%
Send {Ctrl down} {1} {Ctrl Up}

Return
}


GUI_Image_Set (Set_Image) ; options are Stop, Run, Pause, Start
{
   If (Set_image = Stop)
   {
         Guicontrol,hide, Start
         Guicontrol,hide, paused
         Guicontrol,show, Stopped
         Guicontrol,hide, Running        
      }
      
      If (Set_Image = Run)
      {
          Guicontrol,Hide, Start
         Guicontrol,hide, paused
         Guicontrol,Hide, Stopped
         Guicontrol,Show Running     
      }
      
            If (Set_Image = Pause)
      {
          Guicontrol,Hide, Start
         Guicontrol,Show, paused
         Guicontrol,Hide, Stopped
         Guicontrol,hide, Running     
      }
      
            If (Set_Image = Start)
      {
          Guicontrol,Show, Start
         Guicontrol,hide, paused
         Guicontrol,Hide, Stopped
         Guicontrol,hide, Running     
      }
            
       Gui, Submit, NoHide
   return
}
/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Get mouse positions
	THis section does the following:
1. Asks the user to shift click in the approiate locations
2. Records those mouse positiosn and stores them for later use
3. IF not full screen, make acm full screen
4. CLicks in the Prefix box to ensure a proper starting point

	Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the toottip: https://autohotkey.com/docs/commands/ToolTip.htm
	Check the following link for info on the Keywait: https://autohotkey.com/docs/commands/KeyWait.htm
	Check the following link for info on the MouseGetPos: https://autohotkey.com/docs/commands/MouseGetPos.htm
	Check the following link for info on the Click: https://autohotkey.com/docs/commands/Click.htm
	Check the following link for info on the WingetTitle: https://autohotkey.com/docs/commands/WinGetTitle.htm
	
	
	
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
   
   MouseGetPos, Add_Button_X_Location, Add_Button_Y_Location
   
   
   winget,  Active_ID, ID, A 
   ;WinSet, AlwaysOnTop,on,%Title%
   Current_Monitor := GetCurrentMonitor(Active_ID)
   sleep(5)
   tooltip,
   SetTimer,ToolTipTimerbutton,off
   
   sleep(10)
   
   settimer, ToolTipTimerprefix,10
   
   
   Keywait, Shift,D
   Keywait, Lbutton, D							  
   
   MouseGetPos, prefixx, prefixy
   tooltip,
   
   
   settimer, ToolTipTimerprefix,Off
   sleep(10)
   
   
   SetTimer, ToolTipTimerapply, 10  ;timer routine will occur every 10ms..
   
   
   Keywait, Shift,D
   Keywait, Lbutton, D						  
   
   
   MouseGetPos, Applyx, Applyy
   tooltip,
   SetTimer,ToolTipTimerapply,off
   Sleep(10)
   Tooltip, click on effectivity tab
   
   keywait, Shift,D
   Keywait, Lbutton,D
   
   Tooltip
   
   MouseGetPos, Effx,effy
   
   sleep(5)
   ToolTip, 
   Click, %Add_Button_X_Location%, %Add_Button_Y_Location%
   sleep(5)
   Click, %prefixx%, %prefixy%
   sleep(10)
   mousemove, 300,300
   
   Return
}



/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Get_Prefix subroutine
	THis section does the following:
1. Sends controls to the Main Gui window without having it the active window
2. Stores the selected text into a vairable from the main GUI window
3. Swaps some vairables around in case the serial that is entered has multiple eng moldels or not in ACM


	Check the following link for info on the ControlSend:https://autohotkey.com/docs/commands/ControlSend.htm
	Check the following link for info on the ControlGet: https://autohotkey.com/docs/commands/ControlGet.htm
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


Get_Prefix()
{
  global 
  static Prefixcount, PrefixStore1, Prefixstore, Guitextlocation
 
   COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, %Effectivity_Macro% ;note that the controlsend has two commas after the function call (THis always messed me up)
   sleep()
   COntrolSEnd,,{Shift Down}{Right 3}{Shift Up}, %Effectivity_Macro% 
   
   ControlGet,Prefixes,Selected,,,%Effectivity_Macro% 
   
   sleep(3)
   
   If Prefixes = ***
   {
     Prefix_Holder_for_ACM_Input =  complete 
   }
   
   
   COntrolSEnd,, {del}, %Effectivity_Macro% 
   
   Prefix_Holder_for_ACM_Input = %Prefixes% ; I made this instead of just using the Prefixes variable because I could never get the Prefixes variable to hold the value, and this was easier then debugging
   ;Msgbox, Prefix is %Prefix_Holder_for_ACM_Input% ; THis was for debugging purposes
   
   
   Guitextlocation = %Prefixes%
   
   if Prefixcount != 5 ; this is here from so the scrupt knows if it is not first run of the loop. the Prefixcount was set to 5 in the beginning of the script.
     {
	  PrefixStore1 = %PrefixStore%
      Prefixstore =  %Guitextlocation%
      ;msgbox, prefix store is %prefixstore% and prefixstore1 is %PrefixStore1% ;for diag
	 }
   If Prefixcount = 5 ; this If statement only runs during the first run of the loop. 
   {
      PrefixStore1 = %Guitextlocation%
      Prefixstore = %Guitextlocation%
      Prefixcount = 1
      ;msgbox, prefix store is %prefixstore% ; for diag purposes
   }
   
   Return %Prefix_Holder_for_ACM_Input%
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Copy_Serial subroutine
	THis section does the following:
1. Sends controls to the Main Gui window without having it the active window
2. Stores the selected text into a vairable from the main GUI window
3. stores the Serial numbers in a variable depending on if it is the Start or end set of numbers


	Check the following link for info on the ControlSend:https://autohotkey.com/docs/commands/ControlSend.htm
	Check the following link for info on the ControlGet: https://autohotkey.com/docs/commands/ControlGet.htm
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/



Copy_Serial(Prefix_Number_Location_Check)
{
   global
   
   ;~ Prefix_Number_Location_Check++ ; adds +1 to the number in this variable
   
   COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, %Effectivity_Macro%
   sleep()
   COntrolSEnd,,{Shift Down}{Right 5}{Shift Up},%Effectivity_Macro% 
   
   ControlGet, SerialN, Selected,,,%Effectivity_Macro%  ; gets the selected text and stores into the SerialN variable
   
   COntrolSEnd,, {del}, %Effectivity_Macro% 
   
   If Prefix_Number_Location_Check = 1 ; variable will be 1 if it is the first set of Serial numbers (12345-xxxxx)
   {
      ;~ First_Effectivity_Numbers = %SerialN%      
      Prefix_Numbers = %SerialN%      
      
      If Serialzcounter  = 0
      {
         If Serialstore1 = 
         {
            ;~ Serialstore1 = %First_Effectivity_Numbers%
            Serialstore1 = %Prefix_Numbers%
         }
         
         ;~ Serialstore = %First_Effectivity_Numbers%
         Serialstore = %Prefix_Numbers%
         Serialzcounter = 1
      }else  {
         Serialstore1 = %Serialstore%
         ;~ Serialstore =  %First_Effectivity_Numbers%
         Serialstore =  %Prefix_Numbers%
      }
          Addser = %Prefixes%%Prefix_Numbers%
      }
   
   If Prefix_Number_Location_Check = 2 ; variable will be 2 if it is the second set of Serial numbers (xxxxx-12345)
   {
      ;~ Second_Effectivity_Numbers = %SerialN%
      Prefix_Numbers = %SerialN%
      
      If Serialzcounter2  = 0
      {
         If Serialstore3 = 
         {
            ;~ Serialstore3 = %Second_Effectivity_Numbers%
            Serialstore3 = %Prefix_Numbers%
         }
         ;~ Serialstore2 = %Second_Effectivity_Numbers%
         Serialstore2 = %Prefix_Numbers%
         Serialzcounter2 = 1
      }else  {
         Serialstore3 = %Serialstore2%
         ;~ Serialstore2 =  %Second_Effectivity_Numbers%
         Serialstore2 =  %Prefix_Numbers%
      }
      COntrolSEnd,, {del 2}, %Effectivity_Macro% 
      If Prefix_Numbers = 
         Prefix_Numbers = 99999
      
	If Prefix_Numbers = 99999
	Prefix_Numbers = %A_Space%	
    ;~ If Second_Effectivity_Numbers = 99999
	;~ Second_Effectivity_Numbers = %A_Space%
     nextserialtoaddv = %Addser%-%Second_Effectivity_Numbers_Temp%	
      GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
      Gui, submit, nohide
      }
 
   
   Return Prefix_Numbers
}



/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Number_check subroutine
	THis section does the following:
1. Sends controls to the Main Gui window without having it the active window
2. Stores the selected text into a vairable from the main GUI window
3. stores the Serial numbers in a variable depending on if it is the Start or end set of numbers
4.CHecks to see if there is a serial break or if the serial is a x-up and formats it for ACM input


	Check the following link for info on the ControlSend:https://autohotkey.com/docs/commands/ControlSend.htm
	Check the following link for info on the ControlGet: https://autohotkey.com/docs/commands/ControlGet.htm
		Check the following link for info on the GuiControl: https://autohotkey.com/docs/commands/GuiControl.htm
	Check the following link for info on the Gui: https://autohotkey.com/docs/commands/Gui.htm
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


Number_Check()
{
   ;2rd set of numbers check ===========
   COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, %Effectivity_Macro%
   sleep()
   COntrolsend,,{Shift down}{RIght}{SHift up},%Effectivity_Macro%  
   
   ControlGet,Numbers,Selected,,,%Effectivity_Macro% 
   
   COntrolSEnd,, {del},%Effectivity_Macro% 
   
   Serial_Break_Char_Check = %Numbers%
   Return Serial_Break_Char_Check
}


   
Serial_Break_Char(Serial_Break_Char_Check)
{
   ;~ If Serial_Break_Char_Check = -
   ;~ {
      
      ;~ Second_Effectivity_Numbers :=  Copy_Serial(2) ; options are 1  for first set of Prefix numbers, 2 for second set of prefix numbers
      ;~ If Second_Effectivity_Numbers =
      ;~ {
         ;~ Second_Effectivity_Numbers_Temp = 99999
      ;~ }else  {
         ;~ Second_Effectivity_Numbers_Temp = %Second_Effectivity_Numbers%
      ;~ }
      
  
Searchcountser = 0

   
   
   ;~ If Serial_Break_Char_Check = ,
   ;~ {
      ;~ If Prefix_Number_Location_Check = 1
      ;~ {
     
             ;~ Result :=   Searchend()
          ;~ If (Result = Failure) or (Result = Timedout)
            ;~ Exit
         ;~ GOsub, EnterSerials
         ;~ Return
      ;~ }

      
  ;~ thise here -->Do a num check in the correct Location    
  ;~ Else if Prefix_Number_Location_Check = 2  
      ;~ {
         ;~ Addser = %Prefixes%%First_Effectivity_Numbers%
         ;~ nextserialtoaddv = %Addser%-%Second_Effectivity_Numbers%
         ;~ GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
         ;~ COntrolSEnd,, {del},%Effectivity_Macro% 
         ;~ Gui, submit, nohide
       ;~ ;gosub, searchendforserials
  ;~ Result :=   Searchend()
          ;~ If (Result = Failure) or (Result = Timedout)
            ;~ Exit   
         ;~ GOsub, EnterSerials
         ;~ Return
      ;~ }}
   
   ;~ Else if Serial_Break_Char_Check = 
   ;~ {
      ;~ If Prefix_Number_Location_Check = 1
      ;~ {
         ;~ Second_Effectivity_Numbers = %First_Effectivity_Numbers%
         
         ;~ If Second_Effectivity_Numbers =
         ;~ {
            ;~ Second_Effectivity_Numbers_Temp = 99999
         ;~ }else  {
            ;~ Second_Effectivity_Numbers_Temp = %Second_Effectivity_Numbers%
         ;~ }
         
         
      
         ;~ Searchcountser = 0
         ;~ Addser = %Prefixes%%First_Effectivity_Numbers%
         ;~ nextserialtoaddv = %Addser%-%Second_Effectivity_Numbers_Temp%	
               figure this out..  Stopped here.  
               
         if complete = 1
      
         Else
            GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
			
         COntrolSEnd,, {del},%Effectivity_Macro% 
         Gui, submit, nohide
         
           ;~ gosub, searchendforserials	
               Result :=   Searchend()
          If (Result = Failure) or (Result = Timedout)
            Exit
         GOsub, EnterSerials
         Return
      }	
      
      
      Addser = %Prefixes%%First_Effectivity_Numbers%
      nextserialtoaddv = %Addser%-%Second_Effectivity_Numbers_Temp%	
      GuiControl,1:,nextserialtoadd, %nextserialtoaddv%
      COntrolSEnd,, {del},%Effectivity_Macro% 
      Gui, submit, nohide
   }
   
   return
}


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Comma_Check subroutine
	THis section does the following:
1. Sends controls to the Main Gui window without having it the active window
2. Stores the selected text into a vairable from the main GUI window
3. Checks that stored variable to see if it is a comma to find out if the line item in the Main GUi screen is done.


	Check the following link for info on the ControlSend:https://autohotkey.com/docs/commands/ControlSend.htm
	Check the following link for info on the ControlGet: https://autohotkey.com/docs/commands/ControlGet.htm
	
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


Comma_Check(Effectivity_Macro)
{
   COntrolSEnd,, {Shift Down}{Right}{SHift Up},%Effectivity_Macro% 
   ControlGet,Commacheck,Selected,,,%Effectivity_Macro% 
   If Commacheck = ,
   {
      ;Msgbox, COmmafound
      Controlsend,,{BackSpace 2},%Effectivity_Macro% 
   } else  {
      ;msgbox, comma not found
      COntrolSEnd,, {Left}{BackSpace 2}, %Effectivity_Macro% 
      }
   Return
}


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										CHeck_Space subroutine
	THis section does the following:
1. Checks for is there is an empty space in the selected text then removes the space from the GUi screen. This is in case somethign messes up in for formatting.
	- IF i have time, I will check if I even need this here. This may not even be needed anymore as I think it was a workaround for poor formatting of the serials 
	early on when I learning how to do this programming thing..


	Check the following link for info on the ControlSend:https://autohotkey.com/docs/commands/ControlSend.htm

		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


Check_Space:
{
   If First_Effectivity_Numbers = 
   {
      ControlSEnd,,{Del},%Effectivity_Macro% 
      ExitApp
   }
   Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Reloading subroutine
	THis section does the following:
1. Finds the monitor that the mouse is on
2. makes a msgbox to ask user to relaod
3. Sends shift up and control up keys to prevent any hangups if the user reloaded during macro running

	Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


Reloading:
{
     Send {Shift Up}{Ctrl Up} 
   Move_Message_Box("262144", "The number of successful Serial additions to ACM is %Serialcount% `n`n Serial Macro is currently Paused.`n`n Press Okay Button to reload Macro.", "Reload" Effectivity_Macro ) ; THis is to identify the msgbox and move the msgbox to the active screen
   Pause, on
   Reload
   Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Hotkeys declaration section
	THis section does the following:
1. Ensures pressing esc key in only  ACM or in the main gui window will exit the macro 
2. Pressing f1 in the Macro window will bring up the help screen.

	Check the following link for info on the Winactive: https://autohotkey.com/docs/commands/WinActive.htm
		Please note that I seperated the If and the winactive for this so that it can use a vairable for the window name (example... #if Winactive(Variable)
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

#If Winactive (ahk_class TTAFrameXClass)
Esc::
{
   Gosub,Stop_Macro 
   Return
}

#if winactive (Effectivity_Macro)
Esc::
{
   Gosub,Stop_Macro 
   Return
}

F1::
{
   gosub, howto
   Return
}

#if winactive ; stops the requirement for only the macro screen or acm

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										 Stop_macro section
	THis section does the following
1. Finds the monitor that the mouse is on
2. makes a msgbox to ask user to quit
3. If user stops the macro, it does some clean up and stop the program

	Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm
	Check the following link for info on the GuiControl: https://autohotkey.com/docs/commands/GuiControl.htm
	Check the following link for info on the Gui: https://autohotkey.com/docs/commands/Gui.htm
	

\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Stop_Macro:
{
   Msg_box_Result := Move_Message_Box("262148","The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure that you want to stop the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.", "Stop")
   ;~ Msgbox,262148,Stop %Effectivity_Macro%, The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure that you want to stop the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.
   if Msg_box_Result = yes
   {
       tooltip = 
      Stopactcheck = 1						
      Gui 1: -AlwaysOnTop
      GUI_Image_Set ("Stop") ; options are Stop, Run, Pause, Start
      Send {Shift Up}{Ctrl Up}
      breakloop = 1
      Exit
   }
   Else 
      Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										 ExitProgram section
	THis section does the following
1. Finds the monitor that the mouse is on
2. makes a msgbox to ask user to quit
3. If user quits, it does some clean up and exits the program

	Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm
	Check the following link for info on the GuiControl: https://autohotkey.com/docs/commands/GuiControl.htm
	Check the following link for info on the Gui: https://autohotkey.com/docs/commands/Gui.htm
	

\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Exitprogram:
{
 Msg_box_Result :=   Move_Message_Box("262148","The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure that you want to quit the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.","Stop")
   ;~ Msgbox,262148,Stop %Effectivity_Macro%, The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure that you want to quit the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.
      if Msg_box_Result = yes
   {
      tooltip = 
      Stopactcheck = 1						
      Gui 1: -AlwaysOnTop
      GUI_Image_Set ("Stop") ; options are Stop, Run, Pause, Start
      Send {Shift Up}{Ctrl Up}
      breakloop = 1
      Exitapp
   }
   Else 
      Return
}



/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Quit app  \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/
Quitapp:
{
  Msg_box_Result := Move_Message_Box("262148","Are you sure you want to quit","Quit")
   ;~ msgbox,262148,Quit %Effectivity_Macro%, Are you sure you want to quit?
   if Msg_box_Result = yes
   {
       tooltip = 
      Stopactcheck = 1					
      Gui 1: -AlwaysOnTop
      GUI_Image_Set ("Stop") ; options are Stop, Run, Pause, Start
      Send {Shift Up}{Ctrl Up}
      breakloop = 1
      WinSet, AlwaysOnTop, Off,%Title%
      ExitApp
   }else  {
      Return
   }
   Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Win_check subroutine
	THis section does the following:
1. CHecks that the ACM window is the active window. If it is not, then it makes it the active window and waits half a second



	Check the following link for info on the IFWinNotActive: https://autohotkey.com/docs/commands/WinActive.htm
	Check the following link for info on the WinActivate: https://autohotkey.com/docs/commands/WinActivate.htm
	Check the following link for info on the WinWaitActive: https://autohotkey.com/docs/commands/WinWaitActive.htm
	
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Win_check(Active_ID)
{
   IfWinNotActive , ahk_id %Active_ID%
   {
      WinActivate,  ahk_id %Active_ID%
      Sleep()
      WinWaitActive,  ahk_id %Active_ID%,,3
      sleep(5)
   }
   
   return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Searchend subroutine
	THis section does the following:
1. Looks for the Exclimation mark image (Red.png) byt taking a screenshot of the monitor
2. If  it cannot find the image, macro waits half a second and looks again
3. Macro will put a timeout screen after 7 attempts

	Check the following link for info on theListLInes: https://autohotkey.com/docs/commands/ListLines.htm
	Check the following link for info on the Gdip_BitmapFromScreen: Fefer to GDIp library further down this script
		Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm
	Check the following link for info on the Exit: https://autohotkey.com/docs/commands/Exit.htm
	
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


Setup_Images()
{
   global
pToken := Gdip_Startup()
bmpNeedle1 := Gdip_CreateBitmapFromFile(FileNamered)
bmpNeedleSearch := Gdip_CreateBitmapFromFile(FileNameSearch)
bmpNeedleCheck := Gdip_CreateBitmapFromFile(FileNameCheck)
bmpNeedleB := Gdip_CreateBitmapFromFile(FileNameButton)
return pToken
}


Searchend()
{
   global
   ;~ pToken:= Gdip_Startup
   ;~ ;msgbox, searchend
   ;~ listlines off
   ;~ bmpHaystack := Gdip_BitmapFromScreen(Current_Monitor)
   ;~ sleep()
   ;~ RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
   ;~ sleep()
   ;~ ;listlines on
   ;~ If RETSearch < 0
   ;~ {
      ;~ if RETSearch = -1001
      ;~ RETSearch = invalid haystack or needle bitmap pointer
      ;~ if RETSearch = -1002
      ;~ RETSearch = invalid variation value
      ;~ if RETSearch = -1003
      ;~ RETSearch = Unable to lock haystack bitmap bits
      ;~ if RETSearch = -1004
      ;~ RETSearch = Unable to lock needle bitmap bits
      ;~ if RETSearch = -1005
      ;~ RETSearch = Cannot find monitor for screen capture
  ;~ Move_Message_Box("Error Searchend (bmpNeedle1)")
      ;~ Msgbox,262144, %Effectivity_Macro%, Error Searchend (bmpNeedle1) %RetSearch%
      ;~ Exit
   ;~ }
   
   ;~ If RETSearch > 0
   ;~ {
      ;~ SetTimer, refreshcheck, Off
      ;~ Refreshchecks = 0
      ;~ SleepStill = 0 
      ;~ ;Msgbox, found
   ;~ }
   
   
   
   ;~ If RETSearch = 0
   ;~ {
      ;~ If Searchcount = 7
      ;~ {
         ;~ Timeout = Yes
         ;~ gosub, Macrotimedout
      ;~ }
      ;~ ;Msgbox, notfound
      ;~ Sleepstill = 1
      ;~ sleep(5)
      ;~ searchcount++
    ;~ Searchend()
   ;~ }
   ;~ Return
   
     sleep()
   while RETSearch  < 1
   {
   pToken := Setup_Images()
      Current_Monitor := GetCurrentMonitor(Active_ID)
      Sleep()
      bmpHaystack := Gdip_BitmapFromScreen(Current_Monitor)
      sleep()
      RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
     sleep(5)
      Gdip_Shutdown(pToken)
      ;listlines on
      If RETSearch < 0
      {
         Failure :=  RETSearch_Failure_code(RETSearch)
           Move_Message_Box("262144","Error Searchend (bmpNeedle1)" Failure,"Failure")
         ;~ Msgbox,262144, %Effectivity_Macro%, Error Searchend (bmpNeedle1) %Failure%
         Return  "Failure"
         Break         
      }
      
      Attempts_Left := 7 - A_index
            SplashTextOn,,25,Serial Macro, Macro Time will time out after %Attempts_Left% more attemps
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%        

      If A_Index = 7
               {        
               ;~ Macrotimedout()
               Return Timedout
         }
   
}
Return "Success"
}

RETSearch_Failure_code(RETSearch)
{
    if RETSearch = -1001
      Failure = invalid haystack or needle bitmap pointer
      if RETSearch = -1002
      Failure = invalid variation value
      if RETSearch = -1003
      Failure = Unable to lock haystack bitmap bits
      if RETSearch = -1004
      Failure = Unable to lock needle bitmap bits
      if RETSearch = -1005
      Failure = Cannot find monitor for screen capture
      Return Failure
   }
      
;~ }


/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										Searchendforserials subroutine
	THis section does the following:
1. Looks for the Exclimation mark image (Red.png) byt taking a screenshot of the monitor
2. If  it cannot find the image, macro waits half a second and looks again
3. Puts up a timeout screen for the user to show that the macro will time out
4. Macro will put a timeout screen after 7 attempts

	Check the following link for info on theListLInes: https://autohotkey.com/docs/commands/ListLines.htm
	Check the following link for info on the Gdip_BitmapFromScreen: Fefer to GDIp library further down this script
		Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm
	Check the following link for info on the Exit: https://autohotkey.com/docs/commands/Exit.htm
	Check the following link for info on the Splashtext: https://autohotkey.com/docs/commands/SplashTextOn.htm
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Searchendforserials:
{
   
   ;msgbox, end serials
   listlines off
   bmpHaystack := Gdip_BitmapFromScreen(Current_Monitor)
   sleep()
   RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
   sleep()
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
  Move_Message_Box("262144", "Error Searchendserials (bmpNeedle1)" RetSearch, "Failure")
      ;~ Msgbox,262144, %Effectivity_Macro%, Error Searchendserials (bmpNeedle1) %RetSearch%
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
            SplashTextOn,,25,Serial Macro, Macro Time out in 7
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         If Searchcountser = 2
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 6
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         If Searchcountser = 3
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 5
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         If Searchcountser = 4
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 4
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         
         If Searchcountser = 5
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 3
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         If Searchcountser = 6
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 2
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         If Searchcountser = 7
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 1
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         ;listlines on
         If Searchcountser = 8
         {
            SplashTextoff
            Timeout = Yes
            gosub, Macrotimedout
         }}
      
      
      If addtime = 1
      {
         If Searchcountser = 1
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 6
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         If Searchcountser = 2
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 5
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         If Searchcountser = 3
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 4
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         
         If Searchcountser = 4
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 3
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         If Searchcountser = 5
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 2
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         If Searchcountser = 6
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 1
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         ;listlines on
         If Searchcountser = 7
         {
            SplashTextoff
            Timeout = Yes
            gosub, Macrotimedout
         }}else  {
         If Searchcountser = 1
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 5
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         If Searchcountser = 2
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 4
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         
         
         If Searchcountser = 3
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 3
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         If Searchcountser = 4
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 2
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         If Searchcountser = 5
         {
            SplashTextOn,,25,Serial Macro, Macro Time out in 1
            Click, %Applyx%,%Applyy%
            Click, %Applyx%,%Applyy%
         }
         ;listlines on
         If Searchcountser = 6
         {
            SplashTextoff
            Timeout = Yes
            gosub, Macrotimedout
         }}
      
      ;Msgbox, notfound
      Sleepstill = 1
      sleep(5)
      searchcountser++
    ;~ gosub, searchendforserials	
Result :=   Searchend()
          If (Result = Failure) or (Result = Timedout)
            Exit
   }
   Return
}



/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										refreshcheck subroutine
	THis section does the following:
1. adds +1 to the refreshchecks variable
2. If Refreshchecks variablae is equal to 10, make the variuable Needrefresh equal to 1
3. Reset refreshchecks variable to 0
		
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
										GetCurrentMonitor Function
	THis section does the following:
1. Finds the amount of monitors the user has
2. Gets the location of the ACM window (WinGetPOs), and gets the center of that monitor
3. Determines which monitor the ACM window is on


	Check the following link for info on Sysget: https://autohotkey.com/docs/commands/SysGet.htm
	Check the following link for info on wingetpos https://autohotkey.com/docs/commands/WinGetPos.htm

\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

GetCurrentMonitor(Active_ID)
{
   SysGet, numberOfMonitors, MonitorCount
   WinGetPos, winX, winY, winWidth, winHeight, ahk_id %Active_ID%
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
										Refreshpage subroutine
	THis section does the following:
1. Clicks on the add button
2. Goes the the SearchendRefresh subroutine to check if the red Exclimation point is there.
3. Determines which monitor the ACM window is on


	Check the following link for info on Click: https://autohotkey.com/docs/commands/Click.htm

\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

Refreshpage:
{
   sleep(5)
   Click, %Add_Button_X_Location%, %Add_Button_Y_Location%
   sleep(20)
     Result := searchendRefresh()
      If (Result = Failure) 
            Exit   
   Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										SearchendRefresh subroutine
	THis section does the following:
1. Looks for the Exclimation mark image (Red.png) byt taking a screenshot of the monitor after a refresh page operation
2. If  it cannot find the image, macro reruns this subroutine, until the image is found



	Check the following link for info on the Gdip_BitmapFromScreen: Fefer to GDIp library further down this script
	Check the following link for info on the Gdip_ImageSearch: Fefer to GDIp library further down this script
		Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm

		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

SearchendRefresh()
{
   global 
    while RETSearch  < 1
   {
      pToken := Setup_Images()
      Current_Monitor := GetCurrentMonitor(Active_ID)
      Sleep()
      bmpHaystack := Gdip_BitmapFromScreen(Current_Monitor)
      sleep()
      RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,5,0,0,0)
     sleep(5)
      Gdip_Shutdown(pToken)
      ;listlines on
      If RETSearch < 0
      {
         Failure :=  RETSearch_Failure_code(RETSearch)
           Move_Message_Box("262144","Error Searchend (bmpNeedle1)" Failure,"Failure")
         ;~ Msgbox,262144, %Effectivity_Macro%, Error Searchend (bmpNeedle1) %Failure%
         Return  "Failure"
      }}
   
   Return "Success"
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										EnterSerials subroutine
	THis section does the following:
1. Takes the Serial number from the main GUI window and then starts to put it into ACM



	Check the following link for info on the GuiControl: https://autohotkey.com/docs/commands/GuiControl.htm
	Check the following link for info on the Settimer: https://autohotkey.com/docs/commands/SetTimer.htm
	Check the following link for info on the Msgbox: https://autohotkey.com/docs/commands/MsgBox.htm
	Check the following link for info on Click: https://autohotkey.com/docs/commands/Click.htm
	Check the following link for info on MouseMove: https://autohotkey.com/docs/commands/MouseMove.htm
	Check the following link for info on SendRaw: https://autohotkey.com/docs/commands/Send.htm
	Check the following link for info on the Splashtext: https://autohotkey.com/docs/commands/SplashTextOn.htm
		
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/


Enterserials:
{
 Win_check(Active_ID) ; checks to make sure the ACm window is the top one
   
   
   
   Loop, parse, Badlist, `,,all ; Loops through the Bad prefix list so it skips the prefix if the prefix that is about to be entered is already known to not be in ACM
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
   
   Loop, parse, DualENG, `,,all ; Loops though the known Dual Eng prefixes to see if it a know dual prefix
   {
      ;msgbox, loopfield is %A_LoopField%
      
      If prefixes = %A_LoopField%
      {
         ;msgbox,  serial number match
         DUalACMCheck = 1
         DualACMPrefix = %prefixes%
         Break
      }else  {
         ;msgbox, no match
         DUalACMCheck = 0
      }}
   
   
   if Skipserial = 1 ; If skipSerial equals to 1, it adds a modifier to the serial during the Effecticity added screen.
   {
      ;msgbox, no acm skip
      Modifier = -*Serial Not in ACM*
      Guicontrol,1:,Editfield2, %Editfield2%%PrefixStore%%First_Effectivity_Numbers%-%Second_Effectivity_Numbers%%Modifier%`n
      Guicontrol,hide, Editfield2,
      GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%	
      Modifier = 
      Skipserial = 0
       return
   }	
   
   If Complete = 1 ; If the complete variable is equal to 1, the macro stops.
   {
      ElapsedTime := A_TickCount - StartTime
      Total_Time := milli2hms(ElapsedTime, h, m, s)
      sleep(5)
      Send {f5}
       Move_Message_Box("262144", "The number of successful Serial additions to ACM is "  Serialcount  "`n`n Macro Finished due to no more Serials to add. `n`n It took the macro " Total_Time " to perform tasks. `n`n Please close Serial Macro Window when finished checking to ensure serials were entered correctly.","Complete")
      ;~ Msgbox,262144,%Effectivity_Macro%, The number of successful Serial additions to ACM is %Serialcount% `n`n Macro Finished due to no more Serials to add. `n`n It took the macro %Total_Time% to perform tasks. `n`n Please close Serial Macro Window when finished checking to ensure serials were entered correctly.
      Guicontrol,1:, Editfield,
      gosub, radio2h
      GUI_Image_Set ("Stop") ; options are Stop, Run, Pause, Start
       Exit
      Return
   }
   
;---------------------------------------------
; THis starts the part that actually enters the info to the ACM screen
;----------------------------------------------

   ;listlines on
  Result :=   Searchend()
          If (Result = Failure) or (Result = Timedout)
            Exit
   Click, %prefixx%, %prefixy%
   sleep()
   mousemove 300,300
   sleep()
   SEndRaw, %Prefix_Holder_for_ACM_Input%
   sleep()
   Send {Tab}
Win_check(Active_ID)
   Sendraw, %First_Effectivity_Numbers%
   sleep(3)
   Send {Tab}
Win_check(Active_ID)
   SendRaw, %Second_Effectivity_Numbers%
   sleep(3)
   Send {Tab}
   sleep(7)
   If DUalACMCheck = 1
   {
      Modifier = -**Multiple Engineering Models** 
       Move_Message_Box("262144", "This Serial has multiple Engineering Models. Please select the model in ACM and press the OK button in this window.", "Multiple Models")
        ;~ Msgbox,262144,%Effectivity_Macro%, This Serial has multiple Engineering Models. Please select the model in ACM and press the OK button in this window.
      ;Guicontrol,1:,Editfield2, %Editfield2%%Guitextlocation%%First_Effectivity_Numbers%-%Second_Effectivity_Numbers%%Modifier%`n
      SplashTextOn,,25,Serial Macro, Macro will resume in 2
      sleep(10)
      SplashTextOn,,25,Serial Macro, Macro will resume in 1
      sleep(10)
      SplashTextOff
      Win_check(Active_ID)
      Click, %prefixx%, %prefixy%
      sleep(5)
      Send {Tab 3}
      sleep(5)
      DualACMCheck = 0
      SleepStill = 0 
      sleep(3)
   }
   
   IniRead, Sleep_Delay,  C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay
   Sleep(Sleep_Delay)   
   Send {enter 2}
   Sleep(2)
   Click, %Applyx%,%Applyy%
   Click, %Applyx%,%Applyy%
   ;SetTimer, refreshcheck, 250
   Serialcount +=1	
   sleep()
      Searchcountser = 0
   
   
   If Second_Effectivity_Numbers =
   {
      Second_Effectivity_Numbers = 99999
   }
   Guicontrol,1:,Editfield2, %Editfield2%%PrefixStore%%First_Effectivity_Numbers%-%Second_Effectivity_Numbers%%Modifier%`n
   Guicontrol,hide, Editfield2,
   GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%
   ;SetTimer, refreshcheck, Off
   Refreshchecks = 0
   SleepStill = 0 
   Modifier = 
   Skipserial = 0
   Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										milli2hms Function
	THis section does the following:
1. Makes the elaspes time in miliseconds into useable hours, min, and seconds


	Check the following link for info on the Setformat: https://autohotkey.com/docs/commands/SetFormat.htm
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

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
										Exit_label subroutine
	THis section does the following:
1. Turns off the Gdip
2. Exits the macro program
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
										throwout subroutine
	THis section does the following:
1. Makes the bas serial variable equal to one
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/

throwout:
{
   badserial = 1
   Return
}

/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
										ToolTip Timers section
	THis section does the following:
1. all the timer calls for the settimer functions
2. displays the tooltips next to the mouse


	Check the following link for info on the tooltip: https://autohotkey.com/docs/commands/ToolTip.htm
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



Engmodel:
{
   warned = 1
   WinGetPos,ex,ey,,,Macro Timed Out
   IniWrite, %ex%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
   IniWrite, %ey%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition																	
   Gui, 3:Destroy
   Modifier = -**Multiple Engineering Models** 
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
  Move_Message_Box("262144", "Select the engeering model and then press the OK button on this window. The prefix will be logged for this session to prevent more timouts of the same prefix.","Multiple Models")
   ;~ Msgbox,262144, %Effectivity_Macro%, Select the engeering model and then press the OK button on this window. `n`n The prefix will be logged for this session to prevent more timouts of the same prefix.
   sleep()
 Win_check(Active_ID)
   Click, %prefixx%, %prefixy%
   Send {Tab 3}
   SplashTextOn,,25,Serial Macro, Macro will resume in 3
   sleep(3) 
   SplashTextOn,,25,Serial Macro, Macro will resume in 2 
   sleep(3)
   SplashTextOn,,25,Serial Macro, Macro will resume in 1
   sleep(3) 
   
   SplashTextOff
   GUI_Image_Set ("Run") ; options are Stop, Run, Pause, Start
   Pause, Off
   
   Click, %Applyx%,%Applyy%
   Click, %Applyx%,%Applyy%
   Click, %Applyx%,%Applyy%
   modifier = 
Result :=   Searchend()
          If (Result = Failure) or (Result = Timedout)
            Exit
   Return
}

Anarchy:
{
   WinGetPos,ex,ey,,,Macro Timed Out
   IniWrite, %ex%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
   IniWrite, %ey%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition
   Gui, 3:Destroy
  Msg_box_Result := Move_Message_Box("262148","Press Yes button if you found the issue and to Resume the Macro, No button to reload macro", "issues")
   ;~ Msgbox,262148,%Effectivity_Macro%, Press Yes button if you found the issue and to Resume the Macro, No button to reload macro
   If Msg_box_Result =  Yes
   {
      GUI_Image_Set ("Run") ; options are Stop, Run, Pause, Start
       Pause, Off
   }else  {
      GUI_Image_Set ("Stop") ; options are Stop, Run, Pause, Start
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
   Modifier := -*Serial Not in ACM*
   Serialcount--
   ;noacmPrefix = %PrefixStore1%
   ;msgbox, noprefix is %noacmPrefix%
   Badlist = %badlist%%PrefixStore1%`,
   
   ;msgbox, badlist after nogo is %badlist%
   
   SplashTextOn,,25,Serial Macro, Macro will resume in 3
   sleep(5) 
   SplashTextOn,,25,Serial Macro, Macro will resume in 2 
   sleep(5)
   SplashTextOn,,25,Serial Macro, Macro will resume in 1
   sleep(5) 
   SplashTextOff
   
   GUI_Image_Set ("Run") ; options are Stop, Run, Pause, Start
   Pause, Off
Win_check(Active_ID)
   Click, %prefixx%, %prefixy%
   sleep(5)
   Send {ctrl down}{a}{Ctrl up}
   sleep()
   Send {Del}{Tab}
   Send {Del}{Tab}
   Send {Del}{Tab}
   Click, %prefixx%, %prefixy%
   Send {Del}
   
   Modifier = -*Serial Not in ACM*
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
   sleep(20)
   return
}




checkforactivity()
{
     
     while (A_TimeIdlePhysical < 5000)
      {
         Gui 1: -AlwaysOnTop
         activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my )
         
            If Stopactcheck = 1
            {
               SplashTextOff
               Break
            }
            
            If (A_TimeIdlePhysical > 0) and (A_TimeIdlePhysical < 1000)
            Timeleft = 5
        
            If (A_TimeIdlePhysical > 1000) and (A_TimeIdlePhysical < 2000)
            Timeleft = 4
       
            If (A_TimeIdlePhysical > 2000) and (A_TimeIdlePhysical < 3000)
            Timeleft = 3
   
            If (A_TimeIdlePhysical > 3000) and (A_TimeIdlePhysical < 4000)
            Timeleft = 2   
        
            If (A_TimeIdlePhysical > 4000) and (A_TimeIdlePhysical < 5000)
            Timeleft = 1	
            
            SplashTextOn ,350,50,Macro paused, Macro is now paused due to user activity.`n Macro will resume after %timeleft% seconds of no user input
            WinMove, Macro paused,,%amonx%, %Amony%
            pausedstate = yes
            GUI_Image_Set ("Pause") ; options are Stop, Run, Pause, Start
           sleep(10)
         }
         
      if A_TimeIdlePhysical >= 5000 ; meaning there has been no user activity
      {
         If pausedstate = yes
         {
            splashtextoff
            Gui 1: +AlwaysOnTop
         Win_check(Active_ID)
            pausedstate = no
            GUI_Image_Set ("Run") ; options are Stop, Run, Pause, Start
             gosub, radio1h
            Gui, Submit, NoHide
             SerialFullScreen(Active_ID)
            sleep(10)
         }}
         
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
   gosub, pausesub  
   
   Return
   }
   
   pausesub:
   {
      if A_IsPaused = 0
      {
         gosub, radio1h
         Gui 1: -AlwaysOnTop
         GUI_Image_Set ("Pause") ; options are Stop, Run, Pause, Start
         Move_Message_Box("262144","Press pause to unpause",Effectivity_Macro,".1")      
         Move_Message_Box("262144","Press pause to unpause",Effectivity_Macro,".1")      
         Move_Message_Box("262144","Press pause to unpause",Effectivity_Macro,".1")      
         Move_Message_Box("262144","Press pause to unpause",Effectivity_Macro)      
         Pause, toggle, 1 
         Return
      }else  {
         Gui 1: +AlwaysOnTop
      GUI_Image_Set ("Run") ; options are Stop, Run, Pause, Start
        Pause, toggle, 1
      }
      return
   }
   
   SerialFullScreen(Active_ID)
   {
      ;~ WinGetTitle, title, ahk_id %Active_ID%
      WinGetPos,  Xarbor,yarbor,warbor,harbor, ahk_id %Active_ID%
      CurrmonAM := GetCurrentMonitor(Active_ID)
      SysGet,Aarea,MonitorWorkArea,%currmonAM%
      WorkWidthA := AareaRight- AareaLeft
      workHeightA := aareaBottom - aAreaTop
      left := aAreaLeft - 4
      top := AAreaTop - 4
      MouseGetPos mmx,mmy
      If yarbor = %top%
      {
         If xarbor = %left%
         ;Msgbox, win maxed
         Return
      }
      Else
      {
      SysGet,Aarea,Monitor,%currmonAM%
      WidthA := AareaRight- AareaLeft
      HeightA := aareaBottom - aAreaTop
      left := aAreaLeft - 4
      top := AAreaTop - 4
        If yarbor = %top%
         {
         If xarbor = %left%
         ;Msgbox, win maxed
         Return
         } }
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
      GUI_Image_Set ("Run") ; options are Stop, Run, Pause, Start
      GuiControlGet,sleep_delay,,editfield20
      WinGetPos,ex,ey,,,Macro Timed Out
      IniWrite, %ex%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
      IniWrite, %ey%,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition
	  IniWrite, %Sleep_Delay%, C:\Serialmacro\Settings.ini,sleep_delay,sleep_delay
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
         }}
      If addtime = 1
      {
         Addtimetemp++
         If addtimetemp = 2
         {
            Addtime = 2
            Addtimetemp = 0
            Return
         }}
      If addtime = 2
      {
         Addtimetemp++
         If addtimetemp = 2
         {
            Refreshrate /= 2
            Addtimetemp = 0
            addtime = 3
            SplashTextOn,350,50,Refresh time reduced, Macro will refresh ACM every 10 Serials to help reduce the amount of timeouts from a slow ACM
            sleep(50)
            SplashTextOff
            Return
         }}
      If addtime = 3
      {
         Return
      }
      Return
   }
   


/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ FormatSerialMacro Subroutine /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/

FormatSerialsMacro:
{
   ;Clear the variables
   Clipboard5 = 
   FullString = 
   oddballvar = 
   Checkers =
   Clipboard = 
   Clipboard1 = 
   Checkeend  = 
   Editfield = 
   Parseclip = 
   Guicontrol,, Editfield, %Editfield%
   newline = `n
   formatfound = 
   Stringthree = 
   Editfield = 
   FullString = 
   oddballvar = 
   Checkers = 
   clipboard = 
   sleep()
   Send ^c ; sends a control C to the computer to copy selected text
   sleep(2)
   if clipboard =  ; if no text is seleted then clipboard will be remain blank
   {
      Msgbox, Error. Please ensure that text is selected before pressing Ctrl + 1.
      Exit
   }
   FullString = %clipboard%`, ; Sets the variable Fullstring to the clipboard contents
   oddballvar = %Fullstring% ; sets the oddballvar variable to the contents of the fullstring variable
   Checkers = %Fullstring% ; Sets the checkers vairable to the contents of the fullstring variable
   Checkers := Remove_Formatting(Checkers) ; Goes to the Remove_Formatting funciton and stores the completed result into the checkers varialbe
   Clipboard1 := Window_Formatting(Checkers) ; Goes to the Window_Formatting function and stors the completed result into the clipboard1 variable
   ;msgbox, CLipboard1 is `n`n %clipboard1%
   Counter = 0 ; sets the counter variable to 0
   Parseclip = %clipboard1%  ;makes the Parseclip vairable to the contents of the clipboard1 vairable
   Clipboard5 := Parse_Loops(Parseclip) ;  Take info from the Window_Formatting funciton and checks to see of it it just the serial with no numbers attached to it. If is, then adds 00001-99999 to prefix. Also changes the Parseclip variable to the Number of non combined Serials
   
   ;msgbox, CLipboard5 is `n`n %clipboard5%
   Editfieldbreaks = %Clipboard5%%newline% ; makes the editfieldbreaks variable the same as the clipboard5 variable contents and adds a return carriage to it.
  
   CLipboard6 = %Clipboard5% ; store the contents of Clipboard5 into the clipboard6 variable
   ;msgbox, %clipboard6%
   Gosub, Combineserials ;goes to the combine Serials subroutine
   ;Gosub, SerialsGUIscreen
   
   gosub, SerialbreakquestionGUI ; Goes to the Serialsgui.ahk and into the SerialbreakquestionGUI subroutine
   
   ;msgbox, prefixlist is `n`n%Prefixlist% ; FOr diagnostics
   
   Prefixcombinecount = 0 ; Sets the Prefixcombinecount variable to 0
   
   Loop, Parse, Prefixmatching, `, all ; parse loop to breaks the Prefixmatching variable up at the commas
   {
      if a_loopfield =  ; If the text before the comma is nothing, then skip the rest of the loop.
      Continue ; skip over the rest of the loop
      
      
      Prefixcombinecount++ ; Add one to Prefixcombinecount variable
      finalcombine := Prefix%A_LoopField% ; Sets the finalcombine variable to the Prefix variable with the Serial number
      
      If Oneupserial = Yes ; If the 1-UP is selected from the SerialbreakquestionGUI screen, it runs the statement below
      {
         Finalcombine = %A_LoopField%00001-99999 ; Sets the Finalcombine variable to the Prefix variable and adds in 00001-99999
         ;msgbox, %finalcombine%
      }
      ;msgbox, finalcombine is `n %finalcombine% ; For diagnostics
      CLipboard7 = %clipboard7%%finalcombine%`,%newline% ; Sets the clipboard7 variable to Clipboard7 and Finalcombine variable and then adds a comma and a carraige return to make a new line
      
      ;msgbox, matching is `n`n %a_loopfield% ; for diagnostics
   }
   ;msgbox, clipboard7 is `n %clipboard7% ; for diagnostics
   
   EditfieldCombine = %clipboard7% ; Sets the EditfieldCombine to the same as the Clipboard7 variable
   
   If combineser = 0 ; Keep serial Breaks 
   {
      Editfield = %EditFieldbreaks% ; Sets the editfield variable to the Editfieldbreaks variable
      totalprefixes = %Parseclip% ; Sets the totalprefixes variables to the Prefixcombinecount variable
      Guicontrol,1:, Editfield, %EditFieldbreaks% ; Sets the listbox on teh GUi screen to the editfield variable
   }
   
   If combineser = 1 ; combine serial breaks
   {
      Editfield = %EditfieldCombine%%newline% ; Sets the editfield variable to the Editfieldbreaks variable and adds a new line
      totalprefixes = %Prefixcombinecount% ; Sets the totalprefixes variables to the Prefixcombinecount variable
      Guicontrol,1:, Editfield, %EditfieldCombine%%newline% ; Sets the listbox on teh GUi screen to the editfieldcombine vaariable and adds a newline
   }
   
   Guicontrol,, reloadprefixtext, There are a total of %totalprefixes% Serial Numbers to add to ACM ; Changes the valuse in the main GUI screen
   ;sleep(2)
   Winactivate, Effectivity Macro ; Make the Main GUi window  Active
   Guicontrol, Focus, Editfield ; Puts the cursor in the Editfield in teh Gui window
   ;msgbox, pause
   Send {Ctrl Down}{End}{Ctrl Up} ; Send the keystrokes so the cursor is at the bottom of the window
   sleep(3) ; delays the script for 300 miliseconds
   
   Send {BackSpace}{*}{*}{*} ; adds 3 astriks
   send {Ctrl Down}{Home}{Ctrl Up} ; sends keystrokes to move the cursor to the top of the listbox
   gosub, Radio1h ; Goes to the Radio1h subroutine
   
   
   Gui, Submit, NoHide ; Updates the Gui screen
   gosub, ExportSerials
   Return
}

/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ combineserials Subroutine /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/

combineserials:
{
   counting = 0 ; Sets teh vairable to 0
   Match = 0 ; Sets the variable to 0
   
   ;msgbox, clipboard6 is `n`n%clipboard6%
   
   Loop, Parse, clipboard6, `n,all ; loop to divide the clipboard6 variable by the carraige returns
   {
      ;msgbox, combineserials loop is `n%A_LoopField%
      counting++ ; adds 1 to the counting variable
      StringMid, String%Counting%prefix, A_loopfield, 1, 3 ; Gets the prefix letters and stores them for later reuse
	  Checkprefix := String%Counting%prefix ; makes the checkprefix variable the same as the String%Counting%Prefix variable in the line direcly above
      If checkprefix = `, ; checks to see if the checkprefix variable is a comma
      {
         ;msgbox, nothing there
         checkprefix = ; sets the checkprefix variable to nothing
         lastnums =  ; sets the lastnums cariable to nothirng
         Continue ; skips over the rest of the loop and starts at the top of the parse loop
      }
      
      gosub, matchprefix ;goes to the matchprefix subroutine
      
      Stringmid, String%Counting%number,A_LoopField,4,5 ; takes the numbers after teh prefix and stores them into String%Counting%number variable 
      Stringmid, String%Counting%check,A_LoopField,9,1 ; takes the next char after the first half of the serial numbers
      endchar :=  String%Counting%check ; Puts that character into the endchar variable for later use
      begnumcheck := String%Counting%number ; stores the String%Counting%number virable into the begumcheck variable
      ;msgbox, endchar is `n%endchar%
      If endchar = `- ; checks if endchar variable is a hyphen
      {
         Stringmid, String%Counting%last,A_LoopField,10,5 ; IF it was a hypen, then this line grabs the 5 numbers after the hyphen and stores it into String%Counting%last variable
         Lastnums :=  String%Counting%last ; puts the contents of the String%Counting%last variable and puts into lastnums variable for later use
      }
      If Endchar = `, ; if the endchar variable is a comma
      {
         Endchar = `- ; makes the endchar variable a hyphen
         lastnums = %begnumcheck% ; makes the lastnums variable the same as the begumcheck variable
      }
      
      
      ;msgbox, Prefix is `n%Checkprefix% `n`n mischeck is `n%begnumcheck% `n`n checkchar is `n%endchar% `n`n lastnums is `n%lastnums%
      
      If Match = 1 ; If the match variable is 1
      {
	
         Gosub, Checkvalues ; goes to the Checkvalues subroutine
		 
		

		   arrayvar = array%checkprefix% ; FOr testing script
		   Length :=0 ; FOr test script
		   
  for key, value in array%checkprefix% ; FOr testing script 
    Length++ ; FOr testing script
	
		  ; msgbox, % array%checkprefix%.MaxIndex()
		 
	  Length := (Length + 1) ; Test sctipr
		Array%CheckPrefix%[%Length%] := Checkprefix begnumcheck endchar lastnums ; TEst script
		  ;Msgbox, % "matched Length is " . Length . "  Array is " . arrayvar ; FOr testing script
		;msgbox,  % Array%CheckPrefix%[%LEngth%] ; FOr testing script
		
		
         Prefix%checkprefix% = %Checkprefix%%begnumcheck%%endchar%%lastnums%	;makes the Prefix%checkprefix% variable be the combination of the all those other variables
         Serialnumzz := Prefix%Checkprefix% ; makes the Serialnumzz variable the same as the Prefix%Checkprefix% variable
         ;Clipboard7 = %clipboard7%%Serialnumzz%
         Match = 0 ; resets the match variable to 0
         Continue ; stops this iteneration of the loop and starts the next iteneration
      }
      
      Else if Match = 0 ; If the match variable is 0
      {
	     arrayvar = array%checkprefix% ; FOr testing script
		Length :=1 ; FOr testing script
	  
		Array%CheckPrefix%[%Length%] := Checkprefix begnumcheck endchar lastnums ; TEst script
	    ;Msgbox, nomatch `n Length is %Length% `n`n  Array is %arrayvar% ; FOr testing script
		;msgbox,  % Array%CheckPrefix%[%LEngth%] ; FOr testing script
		
		
         Prefix%checkprefix% = %Checkprefix%%begnumcheck%%endchar%%lastnums% ;makes the Prefix%checkprefix% variable be the combination of the all those other variables
      }
      
      ;msgbox, Serialnumzz is `n %Serialnumzz%
      ;msgbox, prefix is  %CHeckprefix%
   }
   
   
   ;msgbox, %Prefixmatching%
   Return
}

/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ Checkvakues Subroutine /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/

Checkvalues:
{
   oldprefix := Prefix%Checkprefix%
   ;msgbox, oldprefix is `n %oldprefix%
   Stringmid, prefixbeg,oldprefix,4,5
   Stringmid, prefixmid,oldprefix,9,1
   endcharcheck :=  prefixmid
   
   Stringmid, Prefixlast,oldprefix,10,5
   Lastnumcheck :=  Prefixlast
   
   
   ;msgbox, oldPrefix is `n%Checkprefix% `n`n prefixbeg is `n%prefixbeg% `n`n endcheckchar is `n%endcharcheck% `n`n prefixlist is `n%Prefixlast%
   

   
   If 	prefixbeg > %lastnums%
   {
      lastnums = %prefixbeg%
   }
   
   If 	prefixbeg < %begnumcheck%
   {
      begnumcheck = %prefixbeg%
   }
   
   
   If 	lastnumcheck > %lastnums%
   {
      lastnums = %lastnumcheck%
   }	
   
   
   return
}

/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ Matchprefix Subroutine /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/

matchprefix:
{
   ;StringReplace, Prefixmatching,prefixmatching,%A_Space%,,all
   Loop, Parse, Prefixmatching,`,,,all
   {
      ;msgbox, loopfield match is %A_LoopField%
      If a_loopfield = 
      {
         blank = 1
         Continue
      }
      
      IF A_LoopField = %CHeckprefix%
      {
	  
         ;msgbox, Match
         Match = 1
		 
         Return
      }}
   
   If match = 1
   {
      return
   }
   
   If blank = 1
   {
      Blank = 0
	}
   
   If match != 1
   {
   
	Array%Checkprefix% := [] ; TEsting code for array
	Blank = 0
    match = 0
    Prefixmatching = %Prefixmatching%%CHeckprefix%`,
   }
   
   return
}

/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\Functions  /\/\//\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ Section /\/\/\\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/
Remove_Formatting(Checkers)
{
   StringReplace, Checkers,Checkers,`n,,All
   StringReplace, Checkers,Checkers,`r,,All
   StringReplace, Checkers,Checkers,`;,`,,All
   StringReplace, Checkers,Checkers,%A_Space%,,All
   StringReplace, Checkers,Checkers, `),`)`n,All
   StringReplace, Checkers,Checkers, 1`-up,,All
   StringReplace, Checkers,Checkers,  `-up,`-99999, All
   StringReplace, Checkers,Checkers, and,,All
   
   Return Checkers
}

Window_Formatting(Checkers)
{
   Counter = 1
   
   ;Loops checkers variable to clean up the all the entereed text. Removes carriage returns with parse, removes any spaces, changes the ) to double , 
   ; Changes 1-up to nothing, changes -up to 999999, changes ), to nothing
   Loop, Parse, Checkers, `r`n
   {
      StringGetPos, pos, A_loopfield, :, 1	
      String%Counter% := SubStr(A_LoopField, pos+2)
      StringTemp := String%Counter%
      ;msgbox, string temp  after sn removal is: `n%Stringtemp%
      StringReplace, newStr,StringTemp,%A_Space%,,All	
      ;msgbox, Newstr after space removal is: `n%Newstr%
      StringReplace, NewStr2,NewStr, `),`,,All
      ;msgbox, Newstr2 after ) to , is: `n %Newstr2%
      StringReplace, NewStr3,NewStr2, 1`-up,,All
      ;msgbox, Newstr3 after 1-up removal is: `n%Newstr3%	
      StringReplace, NewStr4,NewStr3, `-up,`-99999, All
      ;msgbox, Newstr4 after -up to -99999 is: `n %Newstr4%	
      StringReplace, NewStr5,NewStr4, `,,`,`n, All	
      ;	msgbox, Newstr5 after , to ,newline is: `n %Newstr5%
      Counter++
      If Newstr5 = 
      {
         Clipboard1 =  %clipboard1%%NewStr5%
      }else  {
         Clipboard1 =  %clipboard1%%NewStr5%
      }}
   Clipboard1 = %clipboard1%`,
   ;Msgbox, Before replace comma %clipboard1%
   StringReplace, Clipboard1,Clipboard1, `,,,All
   ;StringReplace, Clipboard1,Clipboard1, `,,`)`,,All
   
   ;msgbox, after , to ,) is `n %Clipboard1%
   
   Return Clipboard1
}

Parse_Loops(Byref Parseclip)
{
   TotalPrefixes = 0	
   Counter = 0
   ;Msgbox, parseclip is `n`n %parseclip%
   ;msgbox parseclip is `n%parseclip%
   Loop, parse, Parseclip, `n
   {
      ;msgbox, parseclip loop is `n`n %A_LoopField%
      If a_loopfield = 
      {
         Continue
      }
      TotalPrefixes++	
      Counter++
      addcomma = %A_LoopField%`,
      StringGetPos, pos, addcomma, `,, 1
      If pos = 3
      {
         StringReplace, Clippyt,addcomma,`,,00001`-99999`,`n,all
         ;msgbox, if 3 add 1-up is:`n%Clippyt%
         Clipboard5 = %Clipboard5%%Clippyt%
      }
      Else if pos != 3
      {
         StringMid, String%Counter%prefix, A_loopfield, 1, 3
         Prefixstore := String%Counter%prefix
         ;Prefixlist = %Prefixlist%`,%Prefixstore%
         
         ;msgbox, prefix is %Prefixstore%
         StringTrimLeft, String%Counter%left, A_loopfield, 3
         StringTemppre := String%Counter%left
         ;msgbox, StringTemppre  is %StringTemppre%
         StringTempend := add_digits(StringTemppre)
         ;msgbox, StringTemp  is %StringTempend%
         Stringtemp = %Prefixstore%%Stringtempend%
         StringReplace, StringTemp,StringTemp,`),`,`n,all
         ;msgbox, %stringTemp%
         Clippyt = %Stringtemp%
         ;Msgbox, %A_loopfield% and pos is %pos%
         Clipboard5 = %Clipboard5%%Clippyt%
         ;msgbox,Clipboard5 in loop is %clipboard5%
      }}
   Parseclip = %TotalPrefixes%
   return Clipboard5
}

add_digits(StringTemppre)
{
   Stringtempend =
   combinestring = 
   Loop, Parse, StringTemppre, `-
   {
      StringReplace, StringTemprep,A_LoopField,`),,
      int = %StringTemprep%
      Loop, % 5-StrLen(int)
      int = 0%int%
      combinestring = %combinestring%`-%Int%
   }
   combinestring := SubStr(combinestring, 2)
   Stringtempend = %combinestring%`)
   ;msgbox, end of add zeros is  %Stringtempend%
   
   return Stringtempend
}              



ExportSerials:
{
if checked = 0
Return


Progress, b w200 ,,Creating Excel file (This may take a minute or two)
Progress, 15
;msgbox, start export
GuiControlGet, EditField
;msgbox, %editfield%

Varcount = 1

Loop, parse, Editfield, `,, all
{
var%varcount% := A_LoopField
Varcount++
}

;=== Settings ===
Effectivitycount = 1
WorkbookPath := A_Desktop "\Effectivity.xlsx"    ; full path to your Workbook
Loop
{
IfNotExist %WorkbookPath%
Break 

IfExist %WorkbookPath%
WorkbookPath := A_Desktop "\Effectivity" Effectivitycount ".xlsx"    ; full path to your Workbook
Effectivitycount++
}

SheetName := "Sheet1"            ; name of Worksheet
row := 1                        ; starting row number
column := 1                     ; starting column number


;Progress, b w200 ,,Creating Excel file
Progress, 20
sleep()
oWorkbook := ComObjCreate("Excel.Application") ;handle
;oWorkbook.Visible := True ;by default excel sheets are invisible
oWorkbook.Workbooks.Add ;add a new workbook
Progress, 30
oWorkbook.activeworkbook.SaveAs(WorkbookPath)


MyArray := []   ; make array that will contain values of your variables.
Loop, %Varcount%      ; or more, if you have more variables...
   MyArray.Insert(var%A_Index%)

;Progress, b w200 ,,Creating Excel file
Progress, 50
;=== COM action starts here ===      ; (explaining this could overload you with info, so I won't)

oWorkbook := ComObjGet(WorkbookPath)
oWorkbook.Application.Windows(oWorkbook.Name).Visible := 1
SafeArray := ComObjArray(12, MyArray.MaxIndex(), 1)
Loop, % MyArray.MaxIndex()
SafeArray[A_Index-1, 0] := MyArray[A_Index]
ul := oWorkbook.Worksheets(SheetName).Cells(row, column), lr := oWorkbook.Worksheets(SheetName).Cells(row+MyArray.MaxIndex()-1, column)
oWorkbook.Worksheets(SheetName).Range(ul, lr).Value := SafeArray
;Progress, b w200 ,,Creating Excel file
Progress, 75
oWorkbook.Save
oWorkbook.Close
;Progress, b w200 ,,Creating Excel file
Progress, 100
sleep(5)
Progress, off
;MsgBox, 64,, Done!, 1
;run, %WorkbookPath%
Return
}

CombineStraglers()
{


return
}
 
 /*
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*GUI SEction*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*/ 
aboutmacro:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
			  
   ;msgbox, Prefix is %Guitextlocation%`n Serial is %Second_Effectivity_Numbers%-%Second_Effectivity_Numbers%`n Mod is %Modifier%
   Gui, 7:Add, Picture, x0 y0 w525 h300 +0x4000000, C:\SerialMacro\background.png
   gui, 7:add, Text, xp+5 yp+5 w500 h40 BackgroundTrans, This program was designed to help increase the speed and accuracy of entering machine Effectivity prefixes form the Pubtool into ACM.
   gui, 7:add, Text, xp yp+45 w500 h20 BackgroundTrans, To get the latest version, go to the File menu and select Check for updates.
   gui, 7:add, Text, xp yp+25 w300 h20 BackgroundTrans,  The location of the macro is at the following box account:
   gui 7:font, CBlue Underline
   gui, 7:add, Text, xp+275 yp w500 h20 BackgroundTrans gboxlink, https://cat.box.com/s/eghbsbas6d2qdwsy7y8cyfdglv2331mb
   gui 7:font,
   gui, 7:add, Text, xp-275 yp+25 w500 h40 BackgroundTrans , For reporting bugs or enhancement requests, Please send an email with the Subject line "Effectivity Macro" to the below address
   gui 7:font, CBlue Underline
   gui, 7:add, Text, xp yp+45 w500 h20 BackgroundTrans gemaillink, Karnia_Jarett_S@cat.com
   gui 7:font,
   gui, 7:add, Text, xp+300 yp+45 w500 h20 BackgroundTrans, This program was created by
   gui, 7:add, Text, xp yp+25 w500 h20 BackgroundTrans, and is maintained by Jarett Karnia
   Gui, 7:Show, x%amonx% y%amony% w525 h300, About %Effectivity_Macro%
   Gui, 7: +AlwaysOnTop
   return
}

emaillink:
{
   Run,  mailto:Karnia_Jarett_S@cat?Subject=Effectivity Macro
   return
}

boxlink:
{
   Run, https://cat.box.com/s/eghbsbas6d2qdwsy7y8cyfdglv2331mb
   return
}

SerialbreakquestionGUI:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   gui 1: -alwaysontop
   Gui, 8:Add, Picture, x0 y0 w400 h90 +0x4000000, C:\SerialMacro\background.png
   Gui, 8: Add, text, x10 y20 BackgroundTrans, Do you want to combine the serial breaks, or keep the serial breaks seperated?
   Gui, 8:add, button, xp+50 yp+20 gcombinequstion, Combine
   gui, 8:add, button, xp+75 yp gkeepseperated, Keep Seperated 
   gui, 8:add, button, xp+115 yp goneup, 1-UP all Effectivity
   Gui, 8:Add, Checkbox, XP-190 yp+30 vcreateexcel, Export Excel file of Effectivity to Desktop (Effectivity.xlxs)
   Gui, 8:show, x%amonx% y%amony% w400 h90 
   gui 8: +alwaysontop
   Pausescript()
   return
}

oneup:
{
   UnPausescript()
   GuiControlGet, checked,, createexcel
   Oneupserial = Yes
   Gui, 8:destroy
   combineser = 1
   GOsub, radio1h
   Return
}

combinequstion:
{
   UnPausescript()
   GuiControlGet, checked,, createexcel
   gui 1: +alwaysontop
   Gui, 8:destroy
   combineser = 1
   GOsub, radio1h
   Return
}

keepseperated:
{
   UnPausescript()
   GuiControlGet, checked,, createexcel
   gui 1: +alwaysontop
   Gui, 8:destroy
   combineser = 0
   GOsub, radio1h
   Return
}

SerialsGUIscreen:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
   
   Menu, BBBB, Add, &Check For Update , Versioncheck
   Menu, BBBB, Add, &Options, OptionsGui
   Menu, BBBB, Add, 
   Menu, CCCC, Add, &Run							(Crtl + 2), Enterallserials
   Menu, CCCC, Add, &Pause/Unpause 				(Pause / Insert), pausesub
   Menu, CCCC, Add, &Stop Macro					(ESC), Exitprogram
   Menu, CCCC, Add, &Reload Macro, restartmacro
   Menu, CCCC, Add, &Reload Macro with Current Effectivity, restartmacroEffectivity
   Menu, BBBB, Add, &Exit							(Ctrl + Q), Quitapp
   Menu, DDDD, Add, &How To Use					(F1), HowTo
   Menu, DDDD, Add, &About , Aboutmacro
   
   Menu, MyMenuBar, Add, &File, :BBBB
   Menu, MyMenuBar, Add, &Macro, :CCCC
   Menu, MyMenuBar, Add, &Help, :DDDD
   SplashTextOff
   If totalprefixes < 1
   {
      TotalPrefixzero = 0
   }
   Else 
      TotalPrefixzero = %totalprefixes%
   
   gui 1:add, Edit, x10 y50 w390 h240  vEditField,%editfield%
   
   gui 1:add, Edit, xp yp w390 h240 vEditField2,%editfield2%
   
   Gui 1:Add, Picture, x315 y310 w50 h50 +0x4000000  BackGroundTrans vStarting gstartmacro , C:\SerialMacro\Start.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vRunning, C:\SerialMacro\Running.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vpaused  gpausesub, C:\SerialMacro\Paused.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vStopped grestartmacro, C:\SerialMacro\Stopped.png
   Gui, 1:Add, Picture, x0 y0 w410 h400 +0x4000000 , C:\SerialMacro\background.png
   
   Gui 1:Add, Edit, xp+165 yp+343 w110 h20  vnextserialtoadd, %nextserialtoaddv%
   
   Gui 1:Add, Text, x5 y5 w300 h25 BackgroundTrans +Center vreloadprefixtext, There are a total of %totalprefixzero% Effectivity to add to ACM
   
   Gui 1:add, Radio, xp+25 yp+25 w130 h20 BackGroundTrans vradio1 gradio1h, Effectivity to be added: 
   
   
   
   Gui 1:add, Radio, xp+155 yp w140 h20 BackGroundTrans vradio2 gradio2h, Effectivity already added
   Gui 1:Add, Text, xp-170 Yp+265 W250 h13 BackGroundTrans vserialsentered, Number of Effectivity successfully added to ACM = %Serialcount%
   
   Gui 1:Add, Text, Xp Yp+15 w250 h13  BackGroundTrans , If macro is operating incorrectly, press Esc to reload
   Gui 1:Add, Text, xp yp+15 w250 h13  BackGroundTrans , Or press Pause Button on keyboard to Pause macro, Press Pause again to resume macro.
   Gui 1:Add, Text, xp yp+20 w145 h20  BackgroundTrans , Next Effectivity to add to ACM:
   
   Gui 1:Menu, MyMenuBar
   Gui 1:Show,  x%amonx% y%amony% , %Effectivity_Macro%
   gui 1: +alwaysontop
   Guicontrol,hide, Editfield2,
   Guicontrol,show, Editfield,
   Guicontrol, Focus, Editfield
   IfExist  C:\SerialMacro\Tempcount.txt
   {
      FileRead, Serialcount,C:\SerialMacro\Tempcount.txt
      GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%	
      FileDelete, C:\SerialMacro\Tempcount.txt
   }
   Gui 1:Submit, NoHide
   Return
}

OptionsGui:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   Iniread, refreshrate,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
   Iniread, imagesearchoption,  C:\Serialmacro\Settings.ini,searches,imagesearchoption
   Iniread, Sleep_Delay,  C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay
   gui 10: +alwaysontop
   Gui , 1: -AlwaysOnTop
   ;gui 10:add,radio, checked%mouseclicks% x0 y0 w250 h20 vmouseposclicks gmouseclickss, Mouse click method 
   ;gui 10:add,radio, checked%imagesearchoption% xp yp+25 w317 h20 vimagesearchoption gimagessearching, Image search method `(Not fully functional`)
   gui 10:add, text, x5 y5 w320 h20 ,Refreash ACM Rate (After how many entered effectivity)
   gui 10:add, edit, xp+275 yp-3 w30 veditfield5 , %refreshrate%
   gui 10:add, Text, xp-275 yp+30, ACM speed Compensation (10 = 1 second delay)
   gui 10:add, edit, xp+275 yp-3 w30 veditfield10 , %Sleep_Delay%
   gui 10:add, button, xp-251 yp+26 h20 w75 Default gsavesets, Save Settings
   Gui, 10:Add, Picture, x0 y0 w325 h100 +0x4000000 , C:\SerialMacro\background.png
   gui 10:show, x%amonx% y%amony% w325 h100, Options
   Guicontrol,10:, editfield5, %refreshrate%
   Guicontrol,10:, editfield10, %Sleep_Delay%
   gui, 10:submit, nohide
   return
}

savesets:
{
   Gui , 1: +AlwaysOnTop
   ;msgbox, %refreshrate%
   GuiControlGet,Refreshrate,,editfield5
   GuiControlGet,imagesearchoption
   GuiControlGet,mouseposclicks
   GuiControlGet,sleep_delay,,editfield10
   
   If imagesearchoption = 1
   Entermethod = 1
   
   If mouseposclicks = 1
   Entermethod = 2
   
   IniWrite, %refreshrate%,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
   Iniwrite, %Entermethod%,  C:\Serialmacro\Settings.ini,searches,entermethod
   Iniwrite, %mouseposclicks%,  C:\Serialmacro\Settings.ini,searches,mouseposclicks
   Iniwrite, %imagesearchoption%,  C:\Serialmacro\Settings.ini,searches,imagesearchoption
   IniWrite, %Sleep_Delay%,  C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay
   gui 10:submit, nohide
   gui 10:destroy
   return
}

10guiclose:
{
   Gui , 1: +AlwaysOnTop
   gui 10:destroy
   return
}


mouseclickss:
{
   Entermethod = 2
   Iniwrite, %Entermethod%,  C:\Serialmacro\Settings.ini,searches,entermethod
   return
}

imagessearching:
{
   Entermethod = 1
   Iniwrite, %Entermethod%,  C:\Serialmacro\Settings.ini,searches,entermethod
   return
}


startmacro:
{
   skipbox = 0
   enterallserials()
   ;~ gosub, enterallserials
   return
}

rerunclicks:
{
   Stop = 1
   Send {shift down}{Shift up}
   
   Return
}

pauses:
{
   UnPausescript()
   return
}

radio1h:
{
   GuiControl,, radio2, 0
   GuiControl,, radio1,1
   ;GuiControl,, radio3,0
   ;Storeeditfield2 = %editfield2%`
   Guicontrol,show, Editfield,
   Guicontrol,hide, Editfield2,
   ;Guicontrol,hide, Editfield3,
   gui, submit, nohide
   
   return
}


radio2h:
{
   GuiControl,, radio1,0
   GuiControl,, radio2,1
   ;GuiControl,, radio3,0
   ;Storeeditfield1 = %editfield%
   Guicontrol,show, Editfield2,
   Guicontrol,hide, Editfield,
   ;Guicontrol,hide, Editfield3,
   
   gui, submit, nohide
   ;Guicontrol,, Editfield, %EditField2%
   return
}

radio3h:
{
   GuiControl,, radio1,0
   GuiControl,, radio2,0
   ;GuiControl,, radio3,1
   ;Editfieldtemp = %Editfield%
   ;Editfield = %EditfieldCombine%
   ;Guicontrol,,Editfield,%EditfieldCombine%
   ;Gui, 1:Submit, nohide
   ;Guicontrol,show, Editfield3,
   Guicontrol,hide, Editfield,
   Guicontrol,hide, Editfield2,
   
   ;Guicontrol,, Editfield, %EditField2%
   return
}

Guiclose:
{
Result := Move_Message_Box("262148","Are you sure you want to quit?", "Quit" Effectivity_Macro)
   
   If Result =  Yes
   {
      ;gui destroy
      ExitApp
   }
   
   Return
}

restartmacro:
{
 Msg_box_Result :=  Move_Message_Box("262148","Are you sure that you want to reload the program?", Effectivity_Macro)

   If Msg_box_Result =  yes
   {
      Reload
   }
   return
}


restartmacroEffectivity:
{
Msg_box_Result :=   Move_Message_Box("262148", "Are you sure that you want to reload the program", "reload" Effectivity_Macro)
If Msg_box_Result = yes
   {
      GuiControlGet, nextserialtoadd
      GuiControlGet, EditField
      GuiControlGet, EditField2
      TempSavefile = %nextserialtoadd%`,`n%editfield%      
      FileAppend, %TempSavefile%, C:\SerialMacro\TempAdd.txt 
      FileAppend, %EditField2%, C:\SerialMacro\TempAdded.txt 
      FileAppend, %totalprefixes%, C:\SerialMacro\Tempamount.txt
      FileAppend, %Serialcount%, C:\SerialMacro\Tempcount.txt
      Reload
   }
   return
}


Macrotimedout:
{
   IniRead, amonx,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
   IniRead, amony,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition
   IniRead, Sleep_Delay, C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay
   
   
   If amonx = Error
   {
      activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   }
   
   If Amony = Error
   {
      activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   }
   
   ;msgbox, Prefix is %Guitextlocation%`n Serial is %Second_Effectivity_Numbers%-%Second_Effectivity_Numbers%`n Mod is %Modifier%
   Gui, 3:Add, Picture, x0 y0 w385 h215 +0x4000000, C:\SerialMacro\background.png
   gui, -alwaysontop
   
   
   gui, 3:add, Text, x10 y10 w360  h20 BackgroundTrans,  The delay is most likely due to one of the following:
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - The Eng Model needs to be manually added with the drop down menu.
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - Effectivity # Does not Exist in ACM
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - ACM is running slow
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - Something bad happened and the effectivity page is no more.
   
   gui, 3:add, Text, xp yp+21 w360  h40 BackgroundTrans,  Please press the one of the buttons below after you Idenitify the cause. This action will reflect in the Effectivity Added Screen.
   gui, 3:add, Button, xp+1 yp+45 w80 h50 BackgroundTrans gEngmodel,  More than one Engineering Model
   gui, 3:add, Button, xp+90 yp w80 h50 BackgroundTrans gSerialnogo,  Effectivity Does Not Exist
   gui, 3:add, Button, xp+90 yp w80 h50 BackgroundTrans gacmlong,  ACM Took to long
   gui, 3:add, Button, xp+90 yp w80 h50 BackgroundTrans  gAnarchy,  Something else happened :(
    Gui, 3:add, Text, xp-270 yp+55, ACM Delay (10 = 1 second):
   Gui, 3:add, Edit, xp+140 yp-3 w30 veditfield20 gsavevalue, %Sleep_Delay%
   Gui, 3:add, Text, xp+35 yp+3, <-- Make this Value Higher if Acm is running slow
   Gui, 3:Show, x%amonx% y%amony%, Macro Timed Out %Effectivity_Macro%
   Guicontrol,3:, editfield20, %Sleep_Delay%
   GUI_Image_Set ("Pause") ; options are Stop, Run, Pause, Start
   gui, 3: +alwaysontop
   Pausescript()
   return
}

savevalue:
{
GuiControlGet,Sleep_Delay,,EditField20
IniWrite, %Sleep_Delay%,  C:\Serialmacro\Settings.ini,Sleep_Delay,Sleep_Delay
return
}

Howto:
{
   splashtexton,,Effectivity Macro, Loading PDF
   Run, C:\SerialMacro\How to use Effectivity Macro.pdf
   sleep(20)
   SplashTextOff
   return
}              

 /*
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*Update check section *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 
*/ 

Updatechecker:
{
   if not (FileExist("C:\Serialmacro\Settings.ini"))
   {
      FileAppend,, C:\Serialmacro\Settings.ini
      Msg_box_Result := Move_Message_Box("4"," Looks like This may be the first time running this application. Do you want to check for an update?","Effectivity Macro Updater" )
            if Msg_box_Result = Yes
      {
         gosub, Versioncheck
      }else  {
         IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
         IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
         Return
      }}
   
   if (FileExist("C:\Serialmacro\Settings.ini"))
   {
      IniRead, updatestatus, C:\Serialmacro\Settings.ini, update,lastupdate
      IniRead, Updaterate, C:\Serialmacro\Settings.ini, update,updaterate
      NumberOfDays := A_Now		; Set to the current date first
      EnvSub, NumberOfDays, %updatestatus% , Days 	; this does a date calc, in days
      If NumberOfDays > %Updaterate%	; More than 13 days
      {
        Msg_box_Result :=  Move_Message_Box("4","It has been %NumberOfDays% days since the last update check.`n`n Would you like to check for a new update?", "Effectivity Macro Updater")
      
         if Msg_box_Result = Yes	
         gosub, Versioncheck
         else
         {
            IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
            IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
         }}}
   Return
}


Versioncheck:
{
   Progress,  w200, Updating..., Gathering Information, Effectivity Macro Updater
   Progress, 0
   sleep(2)
   Versioncount = 0
   settimer, versiontimeout, 500
   gosub, create_checkgui
   ;WinSetTitle,,,%a_space%
   Progress,  w200, Updating..., Fetching Server Information, Effectivity Macro Updater
   Progress, 15
   DllCall("SetParent", "uint",  hwnd, "uint", ParentGUI)
   wb.Visible := True
   WinSet, Style, -0xC00000, ahk_id %hwnd%
   
   Progress, 25
   sleep(2)
   
   wb.navigate("https://docs.google.com/document/d/1woiaqcTjqkABrIecRERDAt6nqiEknFWdySqRmie7bCM/edit?usp=sharing")
   
   Progress,  w200, Updating...,Gathering Current Version From Server, Effectivity Macro Updater
   Progress, 50
   
   ;Progress, off
   sleep(2)
   ;msgbox, checking for updates
   ;SplashTextOn , 100, 100, Serial Macros Updater, Checking for update.
   while wb.busy
   {
      sleep()
   }
   
   
   ;ComObjConnect(wb, IE_Events)
   
   Progress,  w200,Updating..., Comparing Version Information, Effectivity Macro Updater
   Progress, 60
   Progress, Off
   splashtextoff
   Winactivate, Serial version
   wingettitle, titleup, A
   ;Msgbox, Title is  %title% 
   StringGetPos, pos, Titleup, #, 1	
   String1 := SubStr(Titleup, pos+2)
   ;	msgbox, %string1%
   ;msgbox, check version is `n`n`n`n%checkversion%
   Checkversion := SubStr(String1,1,3)
   ;msgbox, %Checkversion%
   Progress, Off
   SplashtextOff
   
   If titleup = 
   {
      Winactivate, Serial Version
      wingettitle, titleup, A
      StringGetPos, pos, Titleup, #, 1	
      String1 := SubStr(Titleup, pos+2)
      ;	msgbox, %string1%
      ;msgbox, check version is `n`n`n`n%checkversion%
      ;Checkversion := SubStr(String1,1,5)
   }
   If titleup = 
   {
      Progress,  w200,Updating..., Error Occured. Update Not Able To Complete, Effectivity Macro Updater
      Progress, 0
      sleep(30)
      Progress, off
      SplashTextOff
      msgbox, Error getting update. PLease try again later.
      Gui,2:Destroy
   }
   
   
   If Checkversion >= %Version_Number%
   {
      Gui,2:Destroy
      Progress,  w200,Updating..., Macro is Up to date., Effectivity Macro Updater
      Progress, 100
      sleep(30)
      Progress, off
      SplashTextOff
      settimer, versiontimeout, Off
      IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
      IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
      ;Msgbox,,Serial Macro Updater,Macro is Up to date.
   }
   
   If Checkversion > %Version_Number%
   {
      Gui,2:Destroy			 
      ;gosub, create_checkgui
      
      ;wb.navigate("https://docs.google.com/document/d/1zk6ppZJgpN1vICfJP1VGJDWYpE6XKJXAyKCR5ZH9QTM/edit?usp=sharing")
      ;while wb.busy
      ;{
      ;sleep()
      ;}
      
      
      ;MSGBOX, %clipboard%
      settimer, versiontimeout, Off
Msg_box_Result :=  Move_Message_Box("262148", " New update found. Would you like to open the Cat Box site to download the latest version?", "Effectivity Macro Updater")
  If Msg_box_Result = Yes
      {
         IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
         IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
         Run, https://cat.box.com/s/eghbsbas6d2qdwsy7y8cyfdglv2331mb
      }
      Else
      {
         sleep(5)
      IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
      IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
       }
   Gui,2:Destroy
   return
}}

versiontimeout:
{
   Versioncount++
   If Versioncount = 60
   {
      Progress,  w200,Updating..., Error Occured.Cannot connect to server for update check, please check for internet connection., Effectivity Macro Updater
      Progress, 0
      sleep(30)
      Progress, Off
      SplashTextOff
      ;Msgbox,,Serial Updater Timed out, Cannot connect to server for update check, please check for internet connection.
      Gui,2:Destroy
   }
   Return
}


create_checkgui:
{
   wb := ComObjCreate("InternetExplorer.Application")
   Wb.AddressBar := false
   wb.MenuBar := false
   wb.ToolBar := false
   wb.StatusBar := false
   wb.Resizable := false
   wb.Width := 1000
   wb.Height := 500
   wb.Top := 0
   wb.Left := 0
   wb.Silent := true
   hwnd := wb.hwnd
   Gui, 2:+LastFound
   ParentGUI := WinExist()
   Gui,2:-SysMenu -ToolWindow -Border
   Gui,2:Color, EEAA99                                 
   Gui,2:+LastFound                                     
   WinSet, TransColor, EEAA99                     
   Gui,2:show, x-10 y-10 W1 H1 , Updater
   return
}              
   
   
   activeMonitorInfo( ByRef aX, ByRef aY, ByRef aWidth,  ByRef  aHeight, ByRef mouseX, ByRef mouseY  )
   {
      CoordMode, Mouse, Screen
      MouseGetPos, mouseX , mouseY
      SysGet, monCount, MonitorCount
      Loop %monCount%
      {
         SysGet, curMon, Monitor, %a_index%
         if ( mouseX >= curMonLeft and mouseX <= curMonRight and mouseY >= curMonTop and mouseY <= curMonBottom )
         { 
			aHeight := (curMonBottom - curMonTop)  /2
            aWidth  := (curMonRight  - curMonLeft) /2
            ay     := curMonTop			
            ax      := curMonLeft
			ax 		:= aWidth/2 + ax
			ay		:= aHeight/2 + ay
			;msgbox, ax is %ax% `n ay is %ay% `n aheight is %aHeight% `n awidth is %aWidth%
            return
         }}}
   
   
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
      WinMove, %Msg_box_text% , Amonx, Amony 
      return
}


Move_Message_Box(Msg_box_type, Msg_box_text, Msg_box_title, Msg_box_Time := 2147483 )
{
   global 
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
Settimer, winmovemsgbox, 20
MsgBox, % Msg_box_type , %Msg_box_title% , %Msg_box_text% , %Msg_box_time%
IfMsgBox yes
   Result = Yes
IfMsgBox no
   Result = no

    return Result
}


sleep(Amount := 100)
{
amount := amount * 100
Sleep %Amount%
Return
}

#if

#`:: Listvars

+#`:: ListLines     



















/*
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\. Section of lib files that I got from forums and searching.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
*/
     
;**********************************************************************************
;
; Gdip_ImageSearch()
; by MasterFocus - 02/APRIL/2013 00:30h BRT
; Thanks to guest3456 for helping me ponder some ideas
; Requires GDIP, Gdip_SetBitmapTransColor() and Gdip_MultiLockedBitsSearch()
; http://www.autohotkey.com/board/topic/71100-gdip-imagesearch/
;
; Licensed under CC BY-SA 3.0 -> http://creativecommons.org/licenses/by-sa/3.0/
; I waive compliance with the "Share Alike" condition of the license EXCLUSIVELY
; for these users: tic , Rseding91 , guest3456
;
;==================================================================================
;
; This function searches for pBitmapNeedle within pBitmapHaystack
; The returned value is the number of instances found (negative = error)
;
; ++ PARAMETERS ++
;
; pBitmapHaystack and pBitmapNeedle
;   Self-explanatory bitmap pointers, are the only required parameters
;
; OutputList
;   ByRef variable to store the list of coordinates where a match was found
;
; OuterX1, OuterY1, OuterX2, OuterY2
;   Equivalent to ImageSearch's X1,Y1,X2,Y2
;   Default: 0 for all (which searches the whole haystack area)
;
; Variation
;   Just like ImageSearch, a value from 0 to 255
;   Default: 0
;
; Trans
;   Needle RGB transparent color, should be a numerical value from 0 to 0xFFFFFF
;   Default: blank (does not use transparency)
;
; SearchDirection
;   Haystack search direction
;     Vertical preference:
;       1 = top->left->right->bottom [default]
;       2 = bottom->left->right->top
;       3 = bottom->right->left->top
;       4 = top->right->left->bottom
;     Horizontal preference:
;       5 = left->top->bottom->right
;       6 = left->bottom->top->right
;       7 = right->bottom->top->left
;       8 = right->top->bottom->left
;
; Instances
;   Maximum number of instances to find when searching (0 = find all)
;   Default: 1 (stops after one match is found)
;
; LineDelim and CoordDelim
;   Outer and inner delimiters for the list of coordinates (OutputList)
;   Defaults: "`n" and ","
;
; ++ RETURN VALUES ++
;
; -1001 ==> invalid haystack and/or needle bitmap pointer
; -1002 ==> invalid variation value
; -1003 ==> X1 and Y1 cannot be negative
; -1004 ==> unable to lock haystack bitmap bits
; -1005 ==> unable to lock needle bitmap bits
; any non-negative value ==> the number of instances found
;
;==================================================================================
;
;**********************************************************************************

Gdip_ImageSearch(pBitmapHaystack,pBitmapNeedle,ByRef OutputList=""
,OuterX1=0,OuterY1=0,OuterX2=0,OuterY2=0,Variation=0,Trans=""
,SearchDirection=1,Instances=1,LineDelim="`n",CoordDelim=",") {

    ; Some validations that can be done before proceeding any further
    If !( pBitmapHaystack && pBitmapNeedle )
        Return -1001
    If Variation not between 0 and 255
        return -1002
    If ( ( OuterX1 < 0 ) || ( OuterY1 < 0 ) )
        return -1003
    If SearchDirection not between 1 and 8
        SearchDirection := 1
    If ( Instances < 0 )
        Instances := 0

    ; Getting the dimensions and locking the bits [haystack]
    Gdip_GetImageDimensions(pBitmapHaystack,hWidth,hHeight)
    ; Last parameter being 1 says the LockMode flag is "READ only"
    If Gdip_LockBits(pBitmapHaystack,0,0,hWidth,hHeight,hStride,hScan,hBitmapData,1)
    OR !(hWidth := NumGet(hBitmapData,0))
    OR !(hHeight := NumGet(hBitmapData,4))
        Return -1004

    ; Careful! From this point on, we must do the following before returning:
    ; - unlock haystack bits

    ; Getting the dimensions and locking the bits [needle]
    Gdip_GetImageDimensions(pBitmapNeedle,nWidth,nHeight)
    ; If Trans is correctly specified, create a backup of the original needle bitmap
    ; and modify the current one, setting the desired color as transparent.
    ; Also, since a copy is created, we must remember to dispose the new bitmap later.
    ; This whole thing has to be done before locking the bits.
    If Trans between 0 and 0xFFFFFF
    {
        pOriginalBmpNeedle := pBitmapNeedle
        pBitmapNeedle := Gdip_CloneBitmapArea(pOriginalBmpNeedle,0,0,nWidth,nHeight)
        Gdip_SetBitmapTransColor(pBitmapNeedle,Trans)
        DumpCurrentNeedle := true
    }

    ; Careful! From this point on, we must do the following before returning:
    ; - unlock haystack bits
    ; - dispose current needle bitmap (if necessary)

    If Gdip_LockBits(pBitmapNeedle,0,0,nWidth,nHeight,nStride,nScan,nBitmapData)
    OR !(nWidth := NumGet(nBitmapData,0))
    OR !(nHeight := NumGet(nBitmapData,4))
    {
        If ( DumpCurrentNeedle )
            Gdip_DisposeImage(pBitmapNeedle)
        Gdip_UnlockBits(pBitmapHaystack,hBitmapData)
        Return -1005
    }
    
    ; Careful! From this point on, we must do the following before returning:
    ; - unlock haystack bits
    ; - unlock needle bits
    ; - dispose current needle bitmap (if necessary)

    ; Adjust the search box. "OuterX2,OuterY2" will be the last pixel evaluated
    ; as possibly matching with the needle's first pixel. So, we must avoid going
    ; beyond this maximum final coordinate.
    OuterX2 := ( !OuterX2 ? hWidth-nWidth+1 : OuterX2-nWidth+1 )
    OuterY2 := ( !OuterY2 ? hHeight-nHeight+1 : OuterY2-nHeight+1 )

    OutputCount := Gdip_MultiLockedBitsSearch(hStride,hScan,hWidth,hHeight
    ,nStride,nScan,nWidth,nHeight,OutputList,OuterX1,OuterY1,OuterX2,OuterY2
    ,Variation,SearchDirection,Instances,LineDelim,CoordDelim)

    Gdip_UnlockBits(pBitmapHaystack,hBitmapData)
    Gdip_UnlockBits(pBitmapNeedle,nBitmapData)
    If ( DumpCurrentNeedle )
        Gdip_DisposeImage(pBitmapNeedle)

    Return OutputCount
}

;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;**********************************************************************************
;
; Gdip_SetBitmapTransColor()
; by MasterFocus - 02/APRIL/2013 00:30h BRT
; Requires GDIP
; http://www.autohotkey.com/board/topic/71100-gdip-imagesearch/
;
; Licensed under CC BY-SA 3.0 -> http://creativecommons.org/licenses/by-sa/3.0/
; I waive compliance with the "Share Alike" condition of the license EXCLUSIVELY
; for these users: tic , Rseding91 , guest3456
;
;**********************************************************************************

;==================================================================================
;
; This function modifies the Alpha component for all pixels of a certain color to 0
; The returned value is 0 in case of success, or a negative number otherwise
;
; ++ PARAMETERS ++
;
; pBitmap
;   A valid pointer to the bitmap that will be modified
;
; TransColor
;   The color to become transparent
;   Should range from 0 (black) to 0xFFFFFF (white)
;
; ++ RETURN VALUES ++
;
; -2001 ==> invalid bitmap pointer
; -2002 ==> invalid TransColor
; -2003 ==> unable to retrieve bitmap positive dimensions
; -2004 ==> unable to lock bitmap bits
; -2005 ==> DllCall failed (see ErrorLevel)
; any non-negative value ==> the number of pixels modified by this function
;
;==================================================================================

Gdip_SetBitmapTransColor(pBitmap,TransColor) {
    static _SetBmpTrans, Ptr, PtrA
    if !( _SetBmpTrans ) {
        Ptr := A_PtrSize ? "UPtr" : "UInt"
        PtrA := Ptr . "*"
        MCode_SetBmpTrans := "
            (LTrim Join
            8b44240c558b6c241cc745000000000085c07e77538b5c2410568b74242033c9578b7c2414894c24288da424000000
            0085db7e458bc18d1439b9020000008bff8a0c113a4e0275178a4c38013a4e01750e8a0a3a0e7508c644380300ff450083c0
            0483c204b9020000004b75d38b4c24288b44241c8b5c2418034c242048894c24288944241c75a85f5e5b33c05dc3,405
            34c8b5424388bda41c702000000004585c07e6448897c2410458bd84c8b4424304963f94c8d49010f1f800000000085db7e3
            8498bc1488bd3660f1f440000410fb648023848017519410fb6480138087510410fb6083848ff7507c640020041ff024883c
            00448ffca75d44c03cf49ffcb75bc488b7c241033c05bc3
            )"
        if ( A_PtrSize == 8 ) ; x64, after comma
            MCode_SetBmpTrans := SubStr(MCode_SetBmpTrans,InStr(MCode_SetBmpTrans,",")+1)
        else ; x86, before comma
            MCode_SetBmpTrans := SubStr(MCode_SetBmpTrans,1,InStr(MCode_SetBmpTrans,",")-1)
        VarSetCapacity(_SetBmpTrans, LEN := StrLen(MCode_SetBmpTrans)//2, 0)
        Loop, %LEN%
            NumPut("0x" . SubStr(MCode_SetBmpTrans,(2*A_Index)-1,2), _SetBmpTrans, A_Index-1, "uchar")
        MCode_SetBmpTrans := ""
        DllCall("VirtualProtect", Ptr,&_SetBmpTrans, Ptr,VarSetCapacity(_SetBmpTrans), "uint",0x40, PtrA,0)
    }
    If !pBitmap
        Return -2001
    If TransColor not between 0 and 0xFFFFFF
        Return -2002
    Gdip_GetImageDimensions(pBitmap,W,H)
    If !(W && H)
        Return -2003
    If Gdip_LockBits(pBitmap,0,0,W,H,Stride,Scan,BitmapData)
        Return -2004
    ; The following code should be slower than using the MCode approach,
    ; but will the kept here for now, just for reference.
    /*
    Count := 0
    Loop, %H% {
        Y := A_Index-1
        Loop, %W% {
            X := A_Index-1
            CurrentColor := Gdip_GetLockBitPixel(Scan,X,Y,Stride)
            If ( (CurrentColor & 0xFFFFFF) == TransColor )
                Gdip_SetLockBitPixel(TransColor,Scan,X,Y,Stride), Count++
        }
    }
    */
    ; Thanks guest3456 for helping with the initial solution involving NumPut
    Gdip_FromARGB(TransColor,A,R,G,B), VarSetCapacity(TransColor,0), VarSetCapacity(TransColor,3,255)
    NumPut(B,TransColor,0,"UChar"), NumPut(G,TransColor,1,"UChar"), NumPut(R,TransColor,2,"UChar")
    MCount := 0
    E := DllCall(&_SetBmpTrans, Ptr,Scan, "int",W, "int",H, "int",Stride, Ptr,&TransColor, "int*",MCount, "cdecl int")
    Gdip_UnlockBits(pBitmap,BitmapData)
    If ( E != 0 ) {
        ErrorLevel := E
        Return -2005
    }
    Return MCount
}

;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;**********************************************************************************
;
; Gdip_MultiLockedBitsSearch()
; by MasterFocus - 24/MARCH/2013 06:20h BRT
; Requires GDIP and Gdip_LockedBitsSearch()
; http://www.autohotkey.com/board/topic/71100-gdip-imagesearch/
;
; Licensed under CC BY-SA 3.0 -> http://creativecommons.org/licenses/by-sa/3.0/
; I waive compliance with the "Share Alike" condition of the license EXCLUSIVELY
; for these users: tic , Rseding91 , guest3456
;
;**********************************************************************************

;==================================================================================
;
; This function returns the number of instances found
; The 8 first parameters are the same as in Gdip_LockedBitsSearch()
; The other 10 parameters are the same as in Gdip_ImageSearch()
; Note: the default for the Intances parameter here is 0 (find all matches)
;
;==================================================================================

Gdip_MultiLockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride,nScan,nWidth,nHeight
,ByRef OutputList="",OuterX1=0,OuterY1=0,OuterX2=0,OuterY2=0,Variation=0
,SearchDirection=1,Instances=0,LineDelim="`n",CoordDelim=",")
{
    OutputList := ""
    OutputCount := !Instances
    InnerX1 := OuterX1 , InnerY1 := OuterY1
    InnerX2 := OuterX2 , InnerY2 := OuterY2

    ; The following part is a rather ugly but working hack that I
    ; came up with to adjust the variables and their increments
    ; according to the specified Haystack Search Direction
    /*
    Mod(SD,4) = 0 --> iX = 2 , stepX = +0 , iY = 1 , stepY = +1
    Mod(SD,4) = 1 --> iX = 1 , stepX = +1 , iY = 1 , stepY = +1
    Mod(SD,4) = 2 --> iX = 1 , stepX = +1 , iY = 2 , stepY = +0
    Mod(SD,4) = 3 --> iX = 2 , stepX = +0 , iY = 2 , stepY = +0
    SD <= 4   ------> Vertical preference
    SD > 4    ------> Horizontal preference
    */
    ; Set the index and the step (for both X and Y) to +1
    iX := 1, stepX := 1, iY := 1, stepY := 1
    ; Adjust Y variables if SD is 2, 3, 6 or 7
    Modulo := Mod(SearchDirection,4)
    If ( Modulo > 1 )
        iY := 2, stepY := 0
    ; adjust X variables if SD is 3, 4, 7 or 8
    If !Mod(Modulo,3)
        iX := 2, stepX := 0
    ; Set default Preference to vertical and Nonpreference to horizontal
    P := "Y", N := "X"
    ; adjust Preference and Nonpreference if SD is 5, 6, 7 or 8
    If ( SearchDirection > 4 )
        P := "X", N := "Y"
    ; Set the Preference Index and the Nonpreference Index
    iP := i%P%, iN := i%N%

    While (!(OutputCount == Instances) && (0 == Gdip_LockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride
    ,nScan,nWidth,nHeight,FoundX,FoundY,OuterX1,OuterY1,OuterX2,OuterY2,Variation,SearchDirection)))
    {
        OutputCount++
        OutputList .= LineDelim FoundX CoordDelim FoundY
        Outer%P%%iP% := Found%P%+step%P%
        Inner%N%%iN% := Found%N%+step%N%
        Inner%P%1 := Found%P%
        Inner%P%2 := Found%P%+1
        While (!(OutputCount == Instances) && (0 == Gdip_LockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride
        ,nScan,nWidth,nHeight,FoundX,FoundY,InnerX1,InnerY1,InnerX2,InnerY2,Variation,SearchDirection)))
        {
            OutputCount++
            OutputList .= LineDelim FoundX CoordDelim FoundY
            Inner%N%%iN% := Found%N%+step%N%
        }
    }
    OutputList := SubStr(OutputList,1+StrLen(LineDelim))
    OutputCount -= !Instances
    Return OutputCount
}

;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
;///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

;**********************************************************************************
;
; Gdip_LockedBitsSearch()
; by MasterFocus - 24/MARCH/2013 06:20h BRT
; Mostly adapted from previous work by tic and Rseding91
;
; Requires GDIP
; http://www.autohotkey.com/board/topic/71100-gdip-imagesearch/
;
; Licensed under CC BY-SA 3.0 -> http://creativecommons.org/licenses/by-sa/3.0/
; I waive compliance with the "Share Alike" condition of the license EXCLUSIVELY
; for these users: tic , Rseding91 , guest3456
;
;**********************************************************************************

;==================================================================================
;
; This function searches for a single match of nScan within hScan
;
; ++ PARAMETERS ++
;
; hStride, hScan, hWidth and hHeight
;   Haystack stuff, extracted from a BitmapData, extracted from a Bitmap
;
; nStride, nScan, nWidth and nHeight
;   Needle stuff, extracted from a BitmapData, extracted from a Bitmap
;
; x and y
;   ByRef variables to store the X and Y coordinates of the image if it's found
;   Default: "" for both
;
; sx1, sy1, sx2 and sy2
;   These can be used to crop the search area within the haystack
;   Default: "" for all (does not crop)
;
; Variation
;   Same as the builtin ImageSearch command
;   Default: 0
;
; sd
;   Haystack search direction
;     Vertical preference:
;       1 = top->left->right->bottom [default]
;       2 = bottom->left->right->top
;       3 = bottom->right->left->top
;       4 = top->right->left->bottom
;     Horizontal preference:
;       5 = left->top->bottom->right
;       6 = left->bottom->top->right
;       7 = right->bottom->top->left
;       8 = right->top->bottom->left
;   This value is passed to the internal MCoded function
;
; ++ RETURN VALUES ++
;
; -3001 to -3006 ==> search area incorrectly defined
; -3007 ==> DllCall returned blank
; 0 ==> DllCall succeeded and a match was found
; -4001 ==> DllCall succeeded but a match was not found
; anything else ==> the error value returned by the unsuccessful DllCall
;
;==================================================================================

Gdip_LockedBitsSearch(hStride,hScan,hWidth,hHeight,nStride,nScan,nWidth,nHeight
,ByRef x="",ByRef y="",sx1=0,sy1=0,sx2=0,sy2=0,Variation=0,sd=1)
{
    static _ImageSearch, Ptr, PtrA

    ; Initialize all MCode stuff, if necessary
    if !( _ImageSearch ) {
        Ptr := A_PtrSize ? "UPtr" : "UInt"
        PtrA := Ptr . "*"

        MCode_ImageSearch := "
            (LTrim Join
            8b44243883ec205355565783f8010f857a0100008b7c2458897c24143b7c24600f8db50b00008b44244c8b5c245c8b
            4c24448b7424548be80fafef896c242490897424683bf30f8d0a0100008d64240033c033db8bf5896c241c895c2420894424
            183b4424480f8d0401000033c08944241085c90f8e9d0000008b5424688b7c24408beb8d34968b54246403df8d4900b80300
            0000803c18008b442410745e8b44243c0fb67c2f020fb64c06028d04113bf87f792bca3bf97c738b44243c0fb64c06018b44
            24400fb67c28018d04113bf87f5a2bca3bf97c548b44243c0fb63b0fb60c068d04113bf87f422bca3bf97c3c8b4424108b7c
            24408b4c24444083c50483c30483c604894424103bc17c818b5c24208b74241c0374244c8b44241840035c24508974241ce9
            2dffffff8b6c24688b5c245c8b4c244445896c24683beb8b6c24240f8c06ffffff8b44244c8b7c24148b7424544703e8897c
            2414896c24243b7c24600f8cd5feffffe96b0a00008b4424348b4c246889088b4424388b4c24145f5e5d890833c05b83c420
            c383f8020f85870100008b7c24604f897c24103b7c24580f8c310a00008b44244c8b5c245c8b4c24448bef0fafe8f7d88944
            24188b4424548b742418896c24288d4900894424683bc30f8d0a0100008d64240033c033db8bf5896c2420895c241c894424
            243b4424480f8d0401000033c08944241485c90f8e9d0000008b5424688b7c24408beb8d34968b54246403df8d4900b80300
            0000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f792bca3bf97c738b44243c0fb64c06018b44
            24400fb67c28018d04113bf87f5a2bca3bf97c548b44243c0fb63b0fb60c068d04113bf87f422bca3bf97c3c8b4424148b7c
            24408b4c24444083c50483c30483c604894424143bc17c818b5c241c8b7424200374244c8b44242440035c245089742420e9
            2dffffff8b6c24688b5c245c8b4c244445896c24683beb8b6c24280f8c06ffffff8b7c24108b4424548b7424184f03ee897c
            2410896c24283b7c24580f8dd5feffffe9db0800008b4424348b4c246889088b4424388b4c24105f5e5d890833c05b83c420
            c383f8030f85650100008b7c24604f897c24103b7c24580f8ca10800008b44244c8b6c245c8b5c24548b4c24448bf70faff0
            4df7d8896c242c897424188944241c8bff896c24683beb0f8c020100008d64240033c033db89742424895c2420894424283b
            4424480f8d76ffffff33c08944241485c90f8e9f0000008b5424688b7c24408beb8d34968b54246403dfeb038d4900b80300
            0000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f752bca3bf97c6f8b44243c0fb64c06018b44
            24400fb67c28018d04113bf87f562bca3bf97c508b44243c0fb63b0fb60c068d04113bf87f3e2bca3bf97c388b4424148b7c
            24408b4c24444083c50483c30483c604894424143bc17c818b5c24208b7424248b4424280374244c40035c2450e92bffffff
            8b6c24688b5c24548b4c24448b7424184d896c24683beb0f8d0affffff8b7c24108b44241c4f03f0897c2410897424183b7c
            24580f8c580700008b6c242ce9d4feffff83f8040f85670100008b7c2458897c24103b7c24600f8d340700008b44244c8b6c
            245c8b5c24548b4c24444d8bf00faff7896c242c8974241ceb098da424000000008bff896c24683beb0f8c020100008d6424
            0033c033db89742424895c2420894424283b4424480f8d06feffff33c08944241485c90f8e9f0000008b5424688b7c24408b
            eb8d34968b54246403dfeb038d4900b803000000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f
            752bca3bf97c6f8b44243c0fb64c06018b4424400fb67c28018d04113bf87f562bca3bf97c508b44243c0fb63b0fb60c068d
            04113bf87f3e2bca3bf97c388b4424148b7c24408b4c24444083c50483c30483c604894424143bc17c818b5c24208b742424
            8b4424280374244c40035c2450e92bffffff8b6c24688b5c24548b4c24448b74241c4d896c24683beb0f8d0affffff8b4424
            4c8b7c24104703f0897c24108974241c3b7c24600f8de80500008b6c242ce9d4feffff83f8050f85890100008b7c2454897c
            24683b7c245c0f8dc40500008b5c24608b6c24588b44244c8b4c2444eb078da42400000000896c24103beb0f8d200100008b
            e80faf6c2458896c241c33c033db8bf5896c2424895c2420894424283b4424480f8d0d01000033c08944241485c90f8ea600
            00008b5424688b7c24408beb8d34968b54246403dfeb0a8da424000000008d4900b803000000803c03008b442414745e8b44
            243c0fb67c2f020fb64c06028d04113bf87f792bca3bf97c738b44243c0fb64c06018b4424400fb67c28018d04113bf87f5a
            2bca3bf97c548b44243c0fb63b0fb60c068d04113bf87f422bca3bf97c3c8b4424148b7c24408b4c24444083c50483c30483
            c604894424143bc17c818b5c24208b7424240374244c8b44242840035c245089742424e924ffffff8b7c24108b6c241c8b44
            244c8b5c24608b4c24444703e8897c2410896c241c3bfb0f8cf3feffff8b7c24688b6c245847897c24683b7c245c0f8cc5fe
            ffffe96b0400008b4424348b4c24688b74241089088b4424385f89305e5d33c05b83c420c383f8060f85670100008b7c2454
            897c24683b7c245c0f8d320400008b6c24608b5c24588b44244c8b4c24444d896c24188bff896c24103beb0f8c1a0100008b
            f50faff0f7d88974241c8944242ceb038d490033c033db89742424895c2420894424283b4424480f8d06fbffff33c0894424
            1485c90f8e9f0000008b5424688b7c24408beb8d34968b54246403dfeb038d4900b803000000803c03008b442414745e8b44
            243c0fb67c2f020fb64c06028d04113bf87f752bca3bf97c6f8b44243c0fb64c06018b4424400fb67c28018d04113bf87f56
            2bca3bf97c508b44243c0fb63b0fb60c068d04113bf87f3e2bca3bf97c388b4424148b7c24408b4c24444083c50483c30483
            c604894424143bc17c818b5c24208b7424248b4424280374244c40035c2450e92bffffff8b6c24108b74241c0374242c8b5c
            24588b4c24444d896c24108974241c3beb0f8d02ffffff8b44244c8b7c246847897c24683b7c245c0f8de60200008b6c2418
            e9c2feffff83f8070f85670100008b7c245c4f897c24683b7c24540f8cc10200008b6c24608b5c24588b44244c8b4c24444d
            896c241890896c24103beb0f8c1a0100008bf50faff0f7d88974241c8944242ceb038d490033c033db89742424895c242089
            4424283b4424480f8d96f9ffff33c08944241485c90f8e9f0000008b5424688b7c24408beb8d34968b54246403dfeb038d49
            00b803000000803c18008b442414745e8b44243c0fb67c2f020fb64c06028d04113bf87f752bca3bf97c6f8b44243c0fb64c
            06018b4424400fb67c28018d04113bf87f562bca3bf97c508b44243c0fb63b0fb60c068d04113bf87f3e2bca3bf97c388b44
            24148b7c24408b4c24444083c50483c30483c604894424143bc17c818b5c24208b7424248b4424280374244c40035c2450e9
            2bffffff8b6c24108b74241c0374242c8b5c24588b4c24444d896c24108974241c3beb0f8d02ffffff8b44244c8b7c24684f
            897c24683b7c24540f8c760100008b6c2418e9c2feffff83f8080f85640100008b7c245c4f897c24683b7c24540f8c510100
            008b5c24608b6c24588b44244c8b4c24448d9b00000000896c24103beb0f8d200100008be80faf6c2458896c241c33c033db
            8bf5896c2424895c2420894424283b4424480f8d9dfcffff33c08944241485c90f8ea60000008b5424688b7c24408beb8d34
            968b54246403dfeb0a8da424000000008d4900b803000000803c03008b442414745e8b44243c0fb67c2f020fb64c06028d04
            113bf87f792bca3bf97c738b44243c0fb64c06018b4424400fb67c28018d04113bf87f5a2bca3bf97c548b44243c0fb63b0f
            b604068d0c103bf97f422bc23bf87c3c8b4424148b7c24408b4c24444083c50483c30483c604894424143bc17c818b5c2420
            8b7424240374244c8b44242840035c245089742424e924ffffff8b7c24108b6c241c8b44244c8b5c24608b4c24444703e889
            7c2410896c241c3bfb0f8cf3feffff8b7c24688b6c24584f897c24683b7c24540f8dc5feffff8b4424345fc700ffffffff8b
            4424345e5dc700ffffffffb85ff0ffff5b83c420c3,4c894c24204c89442418488954241048894c24085355565741544
            155415641574883ec188b8424c80000004d8bd94d8bd0488bda83f8010f85b3010000448b8c24a800000044890c24443b8c2
            4b80000000f8d66010000448bac24900000008b9424c0000000448b8424b00000008bbc2480000000448b9424a0000000418
            bcd410fafc9894c24040f1f84000000000044899424c8000000453bd00f8dfb000000468d2495000000000f1f80000000003
            3ed448bf933f6660f1f8400000000003bac24880000000f8d1701000033db85ff7e7e458bf4448bce442bf64503f7904d63c
            14d03c34180780300745a450fb65002438d040e4c63d84c035c2470410fb64b028d0411443bd07f572bca443bd17c50410fb
            64b01450fb650018d0411443bd07f3e2bca443bd17c37410fb60b450fb6108d0411443bd07f272bca443bd17c204c8b5c247
            8ffc34183c1043bdf7c8fffc54503fd03b42498000000e95effffff8b8424c8000000448b8424b00000008b4c24044c8b5c2
            478ffc04183c404898424c8000000413bc00f8c20ffffff448b0c24448b9424a000000041ffc14103cd44890c24894c24044
            43b8c24b80000000f8cd8feffff488b5c2468488b4c2460b85ff0ffffc701ffffffffc703ffffffff4883c418415f415e415
            d415c5f5e5d5bc38b8424c8000000e9860b000083f8020f858c010000448b8c24b800000041ffc944890c24443b8c24a8000
            0007cab448bac2490000000448b8424c00000008b9424b00000008bbc2480000000448b9424a0000000418bc9410fafcd418
            bc5894c2404f7d8894424080f1f400044899424c8000000443bd20f8d02010000468d2495000000000f1f80000000004533f
            6448bf933f60f1f840000000000443bb424880000000f8d56ffffff33db85ff0f8e81000000418bec448bd62bee4103ef496
            3d24903d3807a03007460440fb64a02418d042a4c63d84c035c2470410fb64b02428d0401443bc87f5d412bc8443bc97c554
            10fb64b01440fb64a01428d0401443bc87f42412bc8443bc97c3a410fb60b440fb60a428d0401443bc87f29412bc8443bc97
            c214c8b5c2478ffc34183c2043bdf7c8a41ffc64503fd03b42498000000e955ffffff8b8424c80000008b9424b00000008b4
            c24044c8b5c2478ffc04183c404898424c80000003bc20f8c19ffffff448b0c24448b9424a0000000034c240841ffc9894c2
            40444890c24443b8c24a80000000f8dd0feffffe933feffff83f8030f85c4010000448b8c24b800000041ffc944898c24c80
            00000443b8c24a80000000f8c0efeffff8b842490000000448b9c24b0000000448b8424c00000008bbc248000000041ffcb4
            18bc98bd044895c24080fafc8f7da890c24895424048b9424a0000000448b542404458beb443bda0f8c13010000468d249d0
            000000066660f1f84000000000033ed448bf933f6660f1f8400000000003bac24880000000f8d0801000033db85ff0f8e960
            00000488b4c2478458bf4448bd6442bf64503f70f1f8400000000004963d24803d1807a03007460440fb64a02438d04164c6
            3d84c035c2470410fb64b02428d0401443bc87f63412bc8443bc97c5b410fb64b01440fb64a01428d0401443bc87f48412bc
            8443bc97c40410fb60b440fb60a428d0401443bc87f2f412bc8443bc97c27488b4c2478ffc34183c2043bdf7c8a8b8424900
            00000ffc54403f803b42498000000e942ffffff8b9424a00000008b8424900000008b0c2441ffcd4183ec04443bea0f8d11f
            fffff448b8c24c8000000448b542404448b5c240841ffc94103ca44898c24c8000000890c24443b8c24a80000000f8dc2fef
            fffe983fcffff488b4c24608b8424c8000000448929488b4c2468890133c0e981fcffff83f8040f857f010000448b8c24a80
            0000044890c24443b8c24b80000000f8d48fcffff448bac2490000000448b9424b00000008b9424c0000000448b8424a0000
            0008bbc248000000041ffca418bcd4489542408410fafc9894c2404669044899424c8000000453bd00f8cf8000000468d249
            5000000000f1f800000000033ed448bf933f6660f1f8400000000003bac24880000000f8df7fbffff33db85ff7e7e458bf44
            48bce442bf64503f7904d63c14d03c34180780300745a450fb65002438d040e4c63d84c035c2470410fb64b028d0411443bd
            07f572bca443bd17c50410fb64b01450fb650018d0411443bd07f3e2bca443bd17c37410fb60b450fb6108d0411443bd07f2
            72bca443bd17c204c8b5c2478ffc34183c1043bdf7c8fffc54503fd03b42498000000e95effffff8b8424c8000000448b842
            4a00000008b4c24044c8b5c2478ffc84183ec04898424c8000000413bc00f8d20ffffff448b0c24448b54240841ffc14103c
            d44890c24894c2404443b8c24b80000000f8cdbfeffffe9defaffff83f8050f85ab010000448b8424a000000044890424443
            b8424b00000000f8dc0faffff8b9424c0000000448bac2498000000448ba424900000008bbc2480000000448b8c24a800000
            0428d0c8500000000898c24c800000044894c2404443b8c24b80000000f8d09010000418bc4410fafc18944240833ed448bf
            833f6660f1f8400000000003bac24880000000f8d0501000033db85ff0f8e87000000448bf1448bce442bf64503f74d63c14
            d03c34180780300745d438d040e4c63d84d03da450fb65002410fb64b028d0411443bd07f5f2bca443bd17c58410fb64b014
            50fb650018d0411443bd07f462bca443bd17c3f410fb60b450fb6108d0411443bd07f2f2bca443bd17c284c8b5c24784c8b5
            42470ffc34183c1043bdf7c8c8b8c24c8000000ffc54503fc4103f5e955ffffff448b4424048b4424088b8c24c80000004c8
            b5c24784c8b54247041ffc04103c4448944240489442408443b8424b80000000f8c0effffff448b0424448b8c24a80000004
            1ffc083c10444890424898c24c8000000443b8424b00000000f8cc5feffffe946f9ffff488b4c24608b042489018b4424044
            88b4c2468890133c0e945f9ffff83f8060f85aa010000448b8c24a000000044894c2404443b8c24b00000000f8d0bf9ffff8
            b8424b8000000448b8424c0000000448ba424900000008bbc2480000000428d0c8d00000000ffc88944240c898c24c800000
            06666660f1f840000000000448be83b8424a80000000f8c02010000410fafc4418bd4f7da891424894424084533f6448bf83
            3f60f1f840000000000443bb424880000000f8df900000033db85ff0f8e870000008be9448bd62bee4103ef4963d24903d38
            07a03007460440fb64a02418d042a4c63d84c035c2470410fb64b02428d0401443bc87f64412bc8443bc97c5c410fb64b014
            40fb64a01428d0401443bc87f49412bc8443bc97c41410fb60b440fb60a428d0401443bc87f30412bc8443bc97c284c8b5c2
            478ffc34183c2043bdf7c8a8b8c24c800000041ffc64503fc03b42498000000e94fffffff8b4424088b8c24c80000004c8b5
            c247803042441ffcd89442408443bac24a80000000f8d17ffffff448b4c24048b44240c41ffc183c10444894c2404898c24c
            8000000443b8c24b00000000f8ccefeffffe991f7ffff488b4c24608b4424048901488b4c246833c0448929e992f7ffff83f
            8070f858d010000448b8c24b000000041ffc944894c2404443b8c24a00000000f8c55f7ffff8b8424b8000000448b8424c00
            00000448ba424900000008bbc2480000000428d0c8d00000000ffc8890424898c24c8000000660f1f440000448be83b8424a
            80000000f8c02010000410fafc4418bd4f7da8954240c8944240833ed448bf833f60f1f8400000000003bac24880000000f8
            d4affffff33db85ff0f8e89000000448bf1448bd6442bf64503f74963d24903d3807a03007460440fb64a02438d04164c63d
            84c035c2470410fb64b02428d0401443bc87f63412bc8443bc97c5b410fb64b01440fb64a01428d0401443bc87f48412bc84
            43bc97c40410fb60b440fb60a428d0401443bc87f2f412bc8443bc97c274c8b5c2478ffc34183c2043bdf7c8a8b8c24c8000
            000ffc54503fc03b42498000000e94fffffff8b4424088b8c24c80000004c8b5c24780344240c41ffcd89442408443bac24a
            80000000f8d17ffffff448b4c24048b042441ffc983e90444894c2404898c24c8000000443b8c24a00000000f8dcefeffffe
            9e1f5ffff83f8080f85ddf5ffff448b8424b000000041ffc84489442404443b8424a00000000f8cbff5ffff8b9424c000000
            0448bac2498000000448ba424900000008bbc2480000000448b8c24a8000000428d0c8500000000898c24c800000044890c2
            4443b8c24b80000000f8d08010000418bc4410fafc18944240833ed448bf833f6660f1f8400000000003bac24880000000f8
            d0501000033db85ff0f8e87000000448bf1448bce442bf64503f74d63c14d03c34180780300745d438d040e4c63d84d03da4
            50fb65002410fb64b028d0411443bd07f5f2bca443bd17c58410fb64b01450fb650018d0411443bd07f462bca443bd17c3f4
            10fb603450fb6108d0c10443bd17f2f2bc2443bd07c284c8b5c24784c8b542470ffc34183c1043bdf7c8c8b8c24c8000000f
            fc54503fc4103f5e955ffffff448b04248b4424088b8c24c80000004c8b5c24784c8b54247041ffc04103c44489042489442
            408443b8424b80000000f8c10ffffff448b442404448b8c24a800000041ffc883e9044489442404898c24c8000000443b842
            4a00000000f8dc6feffffe946f4ffff8b442404488b4c246089018b0424488b4c2468890133c0e945f4ffff
            )"
        if ( A_PtrSize == 8 ) ; x64, after comma
            MCode_ImageSearch := SubStr(MCode_ImageSearch,InStr(MCode_ImageSearch,",")+1)
        else ; x86, before comma
            MCode_ImageSearch := SubStr(MCode_ImageSearch,1,InStr(MCode_ImageSearch,",")-1)
        VarSetCapacity(_ImageSearch, LEN := StrLen(MCode_ImageSearch)//2, 0)
        Loop, %LEN%
            NumPut("0x" . SubStr(MCode_ImageSearch,(2*A_Index)-1,2), _ImageSearch, A_Index-1, "uchar")
        MCode_ImageSearch := ""
        DllCall("VirtualProtect", Ptr,&_ImageSearch, Ptr,VarSetCapacity(_ImageSearch), "uint",0x40, PtrA,0)
    }

    ; Abort if an initial coordinates is located before a final coordinate
    If ( sx2 < sx1 )
        return -3001
    If ( sy2 < sy1 )
        return -3002

    ; Check the search box. "sx2,sy2" will be the last pixel evaluated
    ; as possibly matching with the needle's first pixel. So, we must
    ; avoid going beyond this maximum final coordinate.
    If ( sx2 > (hWidth-nWidth+1) )
        return -3003
    If ( sy2 > (hHeight-nHeight+1) )
        return -3004

    ; Abort if the width or height of the search box is 0
    If ( sx2-sx1 == 0 )
        return -3005
    If ( sy2-sy1 == 0 )
        return -3006

    ; The DllCall parameters are the same for easier C code modification,
    ; even though they aren't all used on the _ImageSearch version
    x := 0, y := 0
    , E := DllCall( &_ImageSearch, "int*",x, "int*",y, Ptr,hScan, Ptr,nScan, "int",nWidth, "int",nHeight
    , "int",hStride, "int",nStride, "int",sx1, "int",sy1, "int",sx2, "int",sy2, "int",Variation
    , "int",sd, "cdecl int")
    Return ( E == "" ? -3007 : E )
}

; Gdip standard library v1.45 by tic (Tariq Porter) 07/09/11
; Modifed by Rseding91 using fincs 64 bit compatible Gdip library 5/1/2013
; Supports: Basic, _L ANSi, _L Unicode x86 and _L Unicode x64
;
; Updated 2/20/2014 - fixed Gdip_CreateRegion() and Gdip_GetClipRegion() on AHK Unicode x86
; Updated 5/13/2013 - fixed Gdip_SetBitmapToClipboard() on AHK Unicode x64
;
;#####################################################################################
;#####################################################################################
; STATUS ENUMERATION
; Return values for functions specified to have status enumerated return type
;#####################################################################################
;
; Ok =						= 0
; GenericError				= 1
; InvalidParameter			= 2
; OutOfMemory				= 3
; ObjectBusy				= 4
; InsufficientBuffer		= 5
; NotImplemented			= 6
; Win32Error				= 7
; WrongState				= 8
; Aborted					= 9
; FileNotFound				= 10
; ValueOverflow				= 11
; AccessDenied				= 12
; UnknownImageFormat		= 13
; FontFamilyNotFound		= 14
; FontStyleNotFound			= 15
; NotTrueTypeFont			= 16
; UnsupportedGdiplusVersion	= 17
; GdiplusNotInitialized		= 18
; PropertyNotFound			= 19
; PropertyNotSupported		= 20
; ProfileNotFound			= 21
;
;#####################################################################################
;#####################################################################################
; FUNCTIONS
;#####################################################################################
;
; UpdateLayeredWindow(hwnd, hdc, x="", y="", w="", h="", Alpha=255)
; BitBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, Raster="")
; StretchBlt(dDC, dx, dy, dw, dh, sDC, sx, sy, sw, sh, Raster="")
; SetImage(hwnd, hBitmap)
; Gdip_BitmapFromScreen(Screen=0, Raster="")
; CreateRectF(ByRef RectF, x, y, w, h)
; CreateSizeF(ByRef SizeF, w, h)
; CreateDIBSection
;
;#####################################################################################

; Function:     			UpdateLayeredWindow
; Description:  			Updates a layered window with the handle to the DC of a gdi bitmap
; 
; hwnd        				Handle of the layered window to update
; hdc           			Handle to the DC of the GDI bitmap to update the window with
; Layeredx      			x position to place the window
; Layeredy      			y position to place the window
; Layeredw      			Width of the window
; Layeredh      			Height of the window
; Alpha         			Default = 255 : The transparency (0-255) to set the window transparency
;
; return      				If the function succeeds, the return value is nonzero
;
; notes						If x or y omitted, then layered window will use its current coordinates
;							If w or h omitted then current width and height will be used

UpdateLayeredWindow(hwnd, hdc, x="", y="", w="", h="", Alpha=255)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	if ((x != "") && (y != ""))
		VarSetCapacity(pt, 8), NumPut(x, pt, 0, "UInt"), NumPut(y, pt, 4, "UInt")

	if (w = "") ||(h = "")
		WinGetPos,,, w, h, ahk_id %hwnd%
   
	return DllCall("UpdateLayeredWindow"
					, Ptr, hwnd
					, Ptr, 0
					, Ptr, ((x = "") && (y = "")) ? 0 : &pt
					, "int64*", w|h<<32
					, Ptr, hdc
					, "int64*", 0
					, "uint", 0
					, "UInt*", Alpha<<16|1<<24
					, "uint", 2)
}

;#####################################################################################

; Function				BitBlt
; Description			The BitBlt function performs a bit-block transfer of the color data corresponding to a rectangle 
;						of pixels from the specified source device context into a destination device context.
;
; dDC					handle to destination DC
; dx					x-coord of destination upper-left corner
; dy					y-coord of destination upper-left corner
; dw					width of the area to copy
; dh					height of the area to copy
; sDC					handle to source DC
; sx					x-coordinate of source upper-left corner
; sy					y-coordinate of source upper-left corner
; Raster				raster operation code
;
; return				If the function succeeds, the return value is nonzero
;
; notes					If no raster operation is specified, then SRCCOPY is used, which copies the source directly to the destination rectangle
;
; BLACKNESS				= 0x00000042
; NOTSRCERASE			= 0x001100A6
; NOTSRCCOPY			= 0x00330008
; SRCERASE				= 0x00440328
; DSTINVERT				= 0x00550009
; PATINVERT				= 0x005A0049
; SRCINVERT				= 0x00660046
; SRCAND				= 0x008800C6
; MERGEPAINT			= 0x00BB0226
; MERGECOPY				= 0x00C000CA
; SRCCOPY				= 0x00CC0020
; SRCPAINT				= 0x00EE0086
; PATCOPY				= 0x00F00021
; PATPAINT				= 0x00FB0A09
; WHITENESS				= 0x00FF0062
; CAPTUREBLT			= 0x40000000
; NOMIRRORBITMAP		= 0x80000000

BitBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, Raster="")
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdi32\BitBlt"
					, Ptr, dDC
					, "int", dx
					, "int", dy
					, "int", dw
					, "int", dh
					, Ptr, sDC
					, "int", sx
					, "int", sy
					, "uint", Raster ? Raster : 0x00CC0020)
}

;#####################################################################################

; Function				StretchBlt
; Description			The StretchBlt function copies a bitmap from a source rectangle into a destination rectangle, 
;						stretching or compressing the bitmap to fit the dimensions of the destination rectangle, if necessary.
;						The system stretches or compresses the bitmap according to the stretching mode currently set in the destination device context.
;
; ddc					handle to destination DC
; dx					x-coord of destination upper-left corner
; dy					y-coord of destination upper-left corner
; dw					width of destination rectangle
; dh					height of destination rectangle
; sdc					handle to source DC
; sx					x-coordinate of source upper-left corner
; sy					y-coordinate of source upper-left corner
; sw					width of source rectangle
; sh					height of source rectangle
; Raster				raster operation code
;
; return				If the function succeeds, the return value is nonzero
;
; notes					If no raster operation is specified, then SRCCOPY is used. It uses the same raster operations as BitBlt		

StretchBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, sw, sh, Raster="")
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdi32\StretchBlt"
					, Ptr, ddc
					, "int", dx
					, "int", dy
					, "int", dw
					, "int", dh
					, Ptr, sdc
					, "int", sx
					, "int", sy
					, "int", sw
					, "int", sh
					, "uint", Raster ? Raster : 0x00CC0020)
}

;#####################################################################################

; Function				SetStretchBltMode
; Description			The SetStretchBltMode function sets the bitmap stretching mode in the specified device context
;
; hdc					handle to the DC
; iStretchMode			The stretching mode, describing how the target will be stretched
;
; return				If the function succeeds, the return value is the previous stretching mode. If it fails it will return 0
;
; STRETCH_ANDSCANS 		= 0x01
; STRETCH_ORSCANS 		= 0x02
; STRETCH_DELETESCANS 	= 0x03
; STRETCH_HALFTONE 		= 0x04

SetStretchBltMode(hdc, iStretchMode=4)
{
	return DllCall("gdi32\SetStretchBltMode"
					, A_PtrSize ? "UPtr" : "UInt", hdc
					, "int", iStretchMode)
}

;#####################################################################################

; Function				SetImage
; Description			Associates a new image with a static control
;
; hwnd					handle of the control to update
; hBitmap				a gdi bitmap to associate the static control with
;
; return				If the function succeeds, the return value is nonzero

SetImage(hwnd, hBitmap)
{
	SendMessage, 0x172, 0x0, hBitmap,, ahk_id %hwnd%
	E := ErrorLevel
	DeleteObject(E)
	return E
}

;#####################################################################################

; Function				SetSysColorToControl
; Description			Sets a solid colour to a control
;
; hwnd					handle of the control to update
; SysColor				A system colour to set to the control
;
; return				If the function succeeds, the return value is zero
;
; notes					A control must have the 0xE style set to it so it is recognised as a bitmap
;						By default SysColor=15 is used which is COLOR_3DFACE. This is the standard background for a control
;
; COLOR_3DDKSHADOW				= 21
; COLOR_3DFACE					= 15
; COLOR_3DHIGHLIGHT				= 20
; COLOR_3DHILIGHT				= 20
; COLOR_3DLIGHT					= 22
; COLOR_3DSHADOW				= 16
; COLOR_ACTIVEBORDER			= 10
; COLOR_ACTIVECAPTION			= 2
; COLOR_APPWORKSPACE			= 12
; COLOR_BACKGROUND				= 1
; COLOR_BTNFACE					= 15
; COLOR_BTNHIGHLIGHT			= 20
; COLOR_BTNHILIGHT				= 20
; COLOR_BTNSHADOW				= 16
; COLOR_BTNTEXT					= 18
; COLOR_CAPTIONTEXT				= 9
; COLOR_DESKTOP					= 1
; COLOR_GRADIENTACTIVECAPTION	= 27
; COLOR_GRADIENTINACTIVECAPTION	= 28
; COLOR_GRAYTEXT				= 17
; COLOR_HIGHLIGHT				= 13
; COLOR_HIGHLIGHTTEXT			= 14
; COLOR_HOTLIGHT				= 26
; COLOR_INACTIVEBORDER			= 11
; COLOR_INACTIVECAPTION			= 3
; COLOR_INACTIVECAPTIONTEXT		= 19
; COLOR_INFOBK					= 24
; COLOR_INFOTEXT				= 23
; COLOR_MENU					= 4
; COLOR_MENUHILIGHT				= 29
; COLOR_MENUBAR					= 30
; COLOR_MENUTEXT				= 7
; COLOR_SCROLLBAR				= 0
; COLOR_WINDOW					= 5
; COLOR_WINDOWFRAME				= 6
; COLOR_WINDOWTEXT				= 8

SetSysColorToControl(hwnd, SysColor=15)
{
   WinGetPos,,, w, h, ahk_id %hwnd%
   bc := DllCall("GetSysColor", "Int", SysColor, "UInt")
   pBrushClear := Gdip_BrushCreateSolid(0xff000000 | (bc >> 16 | bc & 0xff00 | (bc & 0xff) << 16))
   pBitmap := Gdip_CreateBitmap(w, h), G := Gdip_GraphicsFromImage(pBitmap)
   Gdip_FillRectangle(G, pBrushClear, 0, 0, w, h)
   hBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap)
   SetImage(hwnd, hBitmap)
   Gdip_DeleteBrush(pBrushClear)
   Gdip_DeleteGraphics(G), Gdip_DisposeImage(pBitmap), DeleteObject(hBitmap)
   return 0
}

;#####################################################################################

; Function				Gdip_BitmapFromScreen
; Description			Gets a gdi+ bitmap from the screen
;
; Screen				0 = All screens
;						Any numerical value = Just that screen
;						x|y|w|h = Take specific coordinates with a width and height
; Raster				raster operation code
;
; return      			If the function succeeds, the return value is a pointer to a gdi+ bitmap
;						-1:		one or more of x,y,w,h not passed properly
;
; notes					If no raster operation is specified, then SRCCOPY is used to the returned bitmap

Gdip_BitmapFromScreen(Screen=0, Raster="")
{
	if (Screen = 0)
	{
		Sysget, x, 76
		Sysget, y, 77	
		Sysget, w, 78
		Sysget, h, 79
	}
	else if (SubStr(Screen, 1, 5) = "hwnd:")
	{
		Screen := SubStr(Screen, 6)
		if !WinExist( "ahk_id " Screen)
			return -2
		WinGetPos,,, w, h, ahk_id %Screen%
		x := y := 0
		hhdc := GetDCEx(Screen, 3)
	}
	else if (Screen&1 != "")
	{
		Sysget, M, Monitor, %Screen%
		x := MLeft, y := MTop, w := MRight-MLeft, h := MBottom-MTop
	}
	else
	{
		StringSplit, S, Screen, |
		x := S1, y := S2, w := S3, h := S4
	}

	if (x = "") || (y = "") || (w = "") || (h = "")
		return -1

	chdc := CreateCompatibleDC(), hbm := CreateDIBSection(w, h, chdc), obm := SelectObject(chdc, hbm), hhdc := hhdc ? hhdc : GetDC()
	BitBlt(chdc, 0, 0, w, h, hhdc, x, y, Raster)
	ReleaseDC(hhdc)
	
	pBitmap := Gdip_CreateBitmapFromHBITMAP(hbm)
	SelectObject(chdc, obm), DeleteObject(hbm), DeleteDC(hhdc), DeleteDC(chdc)
	return pBitmap
}

;#####################################################################################

; Function				Gdip_BitmapFromHWND
; Description			Uses PrintWindow to get a handle to the specified window and return a bitmap from it
;
; hwnd					handle to the window to get a bitmap from
;
; return				If the function succeeds, the return value is a pointer to a gdi+ bitmap
;
; notes					Window must not be not minimised in order to get a handle to it's client area

Gdip_BitmapFromHWND(hwnd)
{
	WinGetPos,,, Width, Height, ahk_id %hwnd%
	hbm := CreateDIBSection(Width, Height), hdc := CreateCompatibleDC(), obm := SelectObject(hdc, hbm)
	PrintWindow(hwnd, hdc)
	pBitmap := Gdip_CreateBitmapFromHBITMAP(hbm)
	SelectObject(hdc, obm), DeleteObject(hbm), DeleteDC(hdc)
	return pBitmap
}

;#####################################################################################

; Function    			CreateRectF
; Description			Creates a RectF object, containing a the coordinates and dimensions of a rectangle
;
; RectF       			Name to call the RectF object
; x            			x-coordinate of the upper left corner of the rectangle
; y            			y-coordinate of the upper left corner of the rectangle
; w            			Width of the rectangle
; h            			Height of the rectangle
;
; return      			No return value

CreateRectF(ByRef RectF, x, y, w, h)
{
   VarSetCapacity(RectF, 16)
   NumPut(x, RectF, 0, "float"), NumPut(y, RectF, 4, "float"), NumPut(w, RectF, 8, "float"), NumPut(h, RectF, 12, "float")
}

;#####################################################################################

; Function    			CreateRect
; Description			Creates a Rect object, containing a the coordinates and dimensions of a rectangle
;
; RectF       			Name to call the RectF object
; x            			x-coordinate of the upper left corner of the rectangle
; y            			y-coordinate of the upper left corner of the rectangle
; w            			Width of the rectangle
; h            			Height of the rectangle
;
; return      			No return value

CreateRect(ByRef Rect, x, y, w, h)
{
	VarSetCapacity(Rect, 16)
	NumPut(x, Rect, 0, "uint"), NumPut(y, Rect, 4, "uint"), NumPut(w, Rect, 8, "uint"), NumPut(h, Rect, 12, "uint")
}
;#####################################################################################

; Function		    	CreateSizeF
; Description			Creates a SizeF object, containing an 2 values
;
; SizeF         		Name to call the SizeF object
; w            			w-value for the SizeF object
; h            			h-value for the SizeF object
;
; return      			No Return value

CreateSizeF(ByRef SizeF, w, h)
{
   VarSetCapacity(SizeF, 8)
   NumPut(w, SizeF, 0, "float"), NumPut(h, SizeF, 4, "float")     
}
;#####################################################################################

; Function		    	CreatePointF
; Description			Creates a SizeF object, containing an 2 values
;
; SizeF         		Name to call the SizeF object
; w            			w-value for the SizeF object
; h            			h-value for the SizeF object
;
; return      			No Return value

CreatePointF(ByRef PointF, x, y)
{
   VarSetCapacity(PointF, 8)
   NumPut(x, PointF, 0, "float"), NumPut(y, PointF, 4, "float")     
}
;#####################################################################################

; Function				CreateDIBSection
; Description			The CreateDIBSection function creates a DIB (Device Independent Bitmap) that applications can write to directly
;
; w						width of the bitmap to create
; h						height of the bitmap to create
; hdc					a handle to the device context to use the palette from
; bpp					bits per pixel (32 = ARGB)
; ppvBits				A pointer to a variable that receives a pointer to the location of the DIB bit values
;
; return				returns a DIB. A gdi bitmap
;
; notes					ppvBits will receive the location of the pixels in the DIB

CreateDIBSection(w, h, hdc="", bpp=32, ByRef ppvBits=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	hdc2 := hdc ? hdc : GetDC()
	VarSetCapacity(bi, 40, 0)
	
	NumPut(w, bi, 4, "uint")
	, NumPut(h, bi, 8, "uint")
	, NumPut(40, bi, 0, "uint")
	, NumPut(1, bi, 12, "ushort")
	, NumPut(0, bi, 16, "uInt")
	, NumPut(bpp, bi, 14, "ushort")
	
	hbm := DllCall("CreateDIBSection"
					, Ptr, hdc2
					, Ptr, &bi
					, "uint", 0
					, A_PtrSize ? "UPtr*" : "uint*", ppvBits
					, Ptr, 0
					, "uint", 0, Ptr)

	if !hdc
		ReleaseDC(hdc2)
	return hbm
}

;#####################################################################################

; Function				PrintWindow
; Description			The PrintWindow function copies a visual window into the specified device context (DC), typically a printer DC
;
; hwnd					A handle to the window that will be copied
; hdc					A handle to the device context
; Flags					Drawing options
;
; return				If the function succeeds, it returns a nonzero value
;
; PW_CLIENTONLY			= 1

PrintWindow(hwnd, hdc, Flags=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("PrintWindow", Ptr, hwnd, Ptr, hdc, "uint", Flags)
}

;#####################################################################################

; Function				DestroyIcon
; Description			Destroys an icon and frees any memory the icon occupied
;
; hIcon					Handle to the icon to be destroyed. The icon must not be in use
;
; return				If the function succeeds, the return value is nonzero

DestroyIcon(hIcon)
{
	return DllCall("DestroyIcon", A_PtrSize ? "UPtr" : "UInt", hIcon)
}

;#####################################################################################

PaintDesktop(hdc)
{
	return DllCall("PaintDesktop", A_PtrSize ? "UPtr" : "UInt", hdc)
}

;#####################################################################################

CreateCompatibleBitmap(hdc, w, h)
{
	return DllCall("gdi32\CreateCompatibleBitmap", A_PtrSize ? "UPtr" : "UInt", hdc, "int", w, "int", h)
}

;#####################################################################################

; Function				CreateCompatibleDC
; Description			This function creates a memory device context (DC) compatible with the specified device
;
; hdc					Handle to an existing device context					
;
; return				returns the handle to a device context or 0 on failure
;
; notes					If this handle is 0 (by default), the function creates a memory device context compatible with the application's current screen

CreateCompatibleDC(hdc=0)
{
   return DllCall("CreateCompatibleDC", A_PtrSize ? "UPtr" : "UInt", hdc)
}

;#####################################################################################

; Function				SelectObject
; Description			The SelectObject function selects an object into the specified device context (DC). The new object replaces the previous object of the same type
;
; hdc					Handle to a DC
; hgdiobj				A handle to the object to be selected into the DC
;
; return				If the selected object is not a region and the function succeeds, the return value is a handle to the object being replaced
;
; notes					The specified object must have been created by using one of the following functions
;						Bitmap - CreateBitmap, CreateBitmapIndirect, CreateCompatibleBitmap, CreateDIBitmap, CreateDIBSection (A single bitmap cannot be selected into more than one DC at the same time)
;						Brush - CreateBrushIndirect, CreateDIBPatternBrush, CreateDIBPatternBrushPt, CreateHatchBrush, CreatePatternBrush, CreateSolidBrush
;						Font - CreateFont, CreateFontIndirect
;						Pen - CreatePen, CreatePenIndirect
;						Region - CombineRgn, CreateEllipticRgn, CreateEllipticRgnIndirect, CreatePolygonRgn, CreateRectRgn, CreateRectRgnIndirect
;
; notes					If the selected object is a region and the function succeeds, the return value is one of the following value
;
; SIMPLEREGION			= 2 Region consists of a single rectangle
; COMPLEXREGION			= 3 Region consists of more than one rectangle
; NULLREGION			= 1 Region is empty

SelectObject(hdc, hgdiobj)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("SelectObject", Ptr, hdc, Ptr, hgdiobj)
}

;#####################################################################################

; Function				DeleteObject
; Description			This function deletes a logical pen, brush, font, bitmap, region, or palette, freeing all system resources associated with the object
;						After the object is deleted, the specified handle is no longer valid
;
; hObject				Handle to a logical pen, brush, font, bitmap, region, or palette to delete
;
; return				Nonzero indicates success. Zero indicates that the specified handle is not valid or that the handle is currently selected into a device context

DeleteObject(hObject)
{
   return DllCall("DeleteObject", A_PtrSize ? "UPtr" : "UInt", hObject)
}

;#####################################################################################

; Function				GetDC
; Description			This function retrieves a handle to a display device context (DC) for the client area of the specified window.
;						The display device context can be used in subsequent graphics display interface (GDI) functions to draw in the client area of the window. 
;
; hwnd					Handle to the window whose device context is to be retrieved. If this value is NULL, GetDC retrieves the device context for the entire screen					
;
; return				The handle the device context for the specified window's client area indicates success. NULL indicates failure

GetDC(hwnd=0)
{
	return DllCall("GetDC", A_PtrSize ? "UPtr" : "UInt", hwnd)
}

;#####################################################################################

; DCX_CACHE = 0x2
; DCX_CLIPCHILDREN = 0x8
; DCX_CLIPSIBLINGS = 0x10
; DCX_EXCLUDERGN = 0x40
; DCX_EXCLUDEUPDATE = 0x100
; DCX_INTERSECTRGN = 0x80
; DCX_INTERSECTUPDATE = 0x200
; DCX_LOCKWINDOWUPDATE = 0x400
; DCX_NORECOMPUTE = 0x100000
; DCX_NORESETATTRS = 0x4
; DCX_PARENTCLIP = 0x20
; DCX_VALIDATE = 0x200000
; DCX_WINDOW = 0x1

GetDCEx(hwnd, flags=0, hrgnClip=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
    return DllCall("GetDCEx", Ptr, hwnd, Ptr, hrgnClip, "int", flags)
}

;#####################################################################################

; Function				ReleaseDC
; Description			This function releases a device context (DC), freeing it for use by other applications. The effect of ReleaseDC depends on the type of device context
;
; hdc					Handle to the device context to be released
; hwnd					Handle to the window whose device context is to be released
;
; return				1 = released
;						0 = not released
;
; notes					The application must call the ReleaseDC function for each call to the GetWindowDC function and for each call to the GetDC function that retrieves a common device context
;						An application cannot use the ReleaseDC function to release a device context that was created by calling the CreateDC function; instead, it must use the DeleteDC function. 

ReleaseDC(hdc, hwnd=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("ReleaseDC", Ptr, hwnd, Ptr, hdc)
}

;#####################################################################################

; Function				DeleteDC
; Description			The DeleteDC function deletes the specified device context (DC)
;
; hdc					A handle to the device context
;
; return				If the function succeeds, the return value is nonzero
;
; notes					An application must not delete a DC whose handle was obtained by calling the GetDC function. Instead, it must call the ReleaseDC function to free the DC

DeleteDC(hdc)
{
   return DllCall("DeleteDC", A_PtrSize ? "UPtr" : "UInt", hdc)
}
;#####################################################################################

; Function				Gdip_LibraryVersion
; Description			Get the current library version
;
; return				the library version
;
; notes					This is useful for non compiled programs to ensure that a person doesn't run an old version when testing your scripts

Gdip_LibraryVersion()
{
	return 1.45
}

;#####################################################################################

; Function				Gdip_LibrarySubVersion
; Description			Get the current library sub version
;
; return				the library sub version
;
; notes					This is the sub-version currently maintained by Rseding91
Gdip_LibrarySubVersion()
{
	return 1.47
}

;#####################################################################################

; Function:    			Gdip_BitmapFromBRA
; Description: 			Gets a pointer to a gdi+ bitmap from a BRA file
;
; BRAFromMemIn			The variable for a BRA file read to memory
; File					The name of the file, or its number that you would like (This depends on alternate parameter)
; Alternate				Changes whether the File parameter is the file name or its number
;
; return      			If the function succeeds, the return value is a pointer to a gdi+ bitmap
;						-1 = The BRA variable is empty
;						-2 = The BRA has an incorrect header
;						-3 = The BRA has information missing
;						-4 = Could not find file inside the BRA

Gdip_BitmapFromBRA(ByRef BRAFromMemIn, File, Alternate=0)
{
	Static FName = "ObjRelease"
	
	if !BRAFromMemIn
		return -1
	Loop, Parse, BRAFromMemIn, `n
	{
		if (A_Index = 1)
		{
			StringSplit, Header, A_LoopField, |
			if (Header0 != 4 || Header2 != "BRA!")
				return -2
		}
		else if (A_Index = 2)
		{
			StringSplit, Info, A_LoopField, |
			if (Info0 != 3)
				return -3
		}
		else
			break
	}
	if !Alternate
		StringReplace, File, File, \, \\, All
	RegExMatch(BRAFromMemIn, "mi`n)^" (Alternate ? File "\|.+?\|(\d+)\|(\d+)" : "\d+\|" File "\|(\d+)\|(\d+)") "$", FileInfo)
	if !FileInfo
		return -4
	
	hData := DllCall("GlobalAlloc", "uint", 2, Ptr, FileInfo2, Ptr)
	pData := DllCall("GlobalLock", Ptr, hData, Ptr)
	DllCall("RtlMoveMemory", Ptr, pData, Ptr, &BRAFromMemIn+Info2+FileInfo1, Ptr, FileInfo2)
	DllCall("GlobalUnlock", Ptr, hData)
	DllCall("ole32\CreateStreamOnHGlobal", Ptr, hData, "int", 1, A_PtrSize ? "UPtr*" : "UInt*", pStream)
	DllCall("gdiplus\GdipCreateBitmapFromStream", Ptr, pStream, A_PtrSize ? "UPtr*" : "UInt*", pBitmap)
	If (A_PtrSize)
		%FName%(pStream)
	Else
		DllCall(NumGet(NumGet(1*pStream)+8), "uint", pStream)
	return pBitmap
}

;#####################################################################################

; Function				Gdip_DrawRectangle
; Description			This function uses a pen to draw the outline of a rectangle into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; x						x-coordinate of the top left of the rectangle
; y						y-coordinate of the top left of the rectangle
; w						width of the rectanlge
; h						height of the rectangle
;
; return				status enumeration. 0 = success
;
; notes					as all coordinates are taken from the top left of each pixel, then the entire width/height should be specified as subtracting the pen width

Gdip_DrawRectangle(pGraphics, pPen, x, y, w, h)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipDrawRectangle", Ptr, pGraphics, Ptr, pPen, "float", x, "float", y, "float", w, "float", h)
}

;#####################################################################################

; Function				Gdip_DrawRoundedRectangle
; Description			This function uses a pen to draw the outline of a rounded rectangle into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; x						x-coordinate of the top left of the rounded rectangle
; y						y-coordinate of the top left of the rounded rectangle
; w						width of the rectanlge
; h						height of the rectangle
; r						radius of the rounded corners
;
; return				status enumeration. 0 = success
;
; notes					as all coordinates are taken from the top left of each pixel, then the entire width/height should be specified as subtracting the pen width

Gdip_DrawRoundedRectangle(pGraphics, pPen, x, y, w, h, r)
{
	Gdip_SetClipRect(pGraphics, x-r, y-r, 2*r, 2*r, 4)
	Gdip_SetClipRect(pGraphics, x+w-r, y-r, 2*r, 2*r, 4)
	Gdip_SetClipRect(pGraphics, x-r, y+h-r, 2*r, 2*r, 4)
	Gdip_SetClipRect(pGraphics, x+w-r, y+h-r, 2*r, 2*r, 4)
	E := Gdip_DrawRectangle(pGraphics, pPen, x, y, w, h)
	Gdip_ResetClip(pGraphics)
	Gdip_SetClipRect(pGraphics, x-(2*r), y+r, w+(4*r), h-(2*r), 4)
	Gdip_SetClipRect(pGraphics, x+r, y-(2*r), w-(2*r), h+(4*r), 4)
	Gdip_DrawEllipse(pGraphics, pPen, x, y, 2*r, 2*r)
	Gdip_DrawEllipse(pGraphics, pPen, x+w-(2*r), y, 2*r, 2*r)
	Gdip_DrawEllipse(pGraphics, pPen, x, y+h-(2*r), 2*r, 2*r)
	Gdip_DrawEllipse(pGraphics, pPen, x+w-(2*r), y+h-(2*r), 2*r, 2*r)
	Gdip_ResetClip(pGraphics)
	return E
}

;#####################################################################################

; Function				Gdip_DrawEllipse
; Description			This function uses a pen to draw the outline of an ellipse into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; x						x-coordinate of the top left of the rectangle the ellipse will be drawn into
; y						y-coordinate of the top left of the rectangle the ellipse will be drawn into
; w						width of the ellipse
; h						height of the ellipse
;
; return				status enumeration. 0 = success
;
; notes					as all coordinates are taken from the top left of each pixel, then the entire width/height should be specified as subtracting the pen width

Gdip_DrawEllipse(pGraphics, pPen, x, y, w, h)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipDrawEllipse", Ptr, pGraphics, Ptr, pPen, "float", x, "float", y, "float", w, "float", h)
}

;#####################################################################################

; Function				Gdip_DrawBezier
; Description			This function uses a pen to draw the outline of a bezier (a weighted curve) into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; x1					x-coordinate of the start of the bezier
; y1					y-coordinate of the start of the bezier
; x2					x-coordinate of the first arc of the bezier
; y2					y-coordinate of the first arc of the bezier
; x3					x-coordinate of the second arc of the bezier
; y3					y-coordinate of the second arc of the bezier
; x4					x-coordinate of the end of the bezier
; y4					y-coordinate of the end of the bezier
;
; return				status enumeration. 0 = success
;
; notes					as all coordinates are taken from the top left of each pixel, then the entire width/height should be specified as subtracting the pen width

Gdip_DrawBezier(pGraphics, pPen, x1, y1, x2, y2, x3, y3, x4, y4)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipDrawBezier"
					, Ptr, pgraphics
					, Ptr, pPen
					, "float", x1
					, "float", y1
					, "float", x2
					, "float", y2
					, "float", x3
					, "float", y3
					, "float", x4
					, "float", y4)
}

;#####################################################################################

; Function				Gdip_DrawArc
; Description			This function uses a pen to draw the outline of an arc into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; x						x-coordinate of the start of the arc
; y						y-coordinate of the start of the arc
; w						width of the arc
; h						height of the arc
; StartAngle			specifies the angle between the x-axis and the starting point of the arc
; SweepAngle			specifies the angle between the starting and ending points of the arc
;
; return				status enumeration. 0 = success
;
; notes					as all coordinates are taken from the top left of each pixel, then the entire width/height should be specified as subtracting the pen width

Gdip_DrawArc(pGraphics, pPen, x, y, w, h, StartAngle, SweepAngle)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipDrawArc"
					, Ptr, pGraphics
					, Ptr, pPen
					, "float", x
					, "float", y
					, "float", w
					, "float", h
					, "float", StartAngle
					, "float", SweepAngle)
}

;#####################################################################################

; Function				Gdip_DrawPie
; Description			This function uses a pen to draw the outline of a pie into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; x						x-coordinate of the start of the pie
; y						y-coordinate of the start of the pie
; w						width of the pie
; h						height of the pie
; StartAngle			specifies the angle between the x-axis and the starting point of the pie
; SweepAngle			specifies the angle between the starting and ending points of the pie
;
; return				status enumeration. 0 = success
;
; notes					as all coordinates are taken from the top left of each pixel, then the entire width/height should be specified as subtracting the pen width

Gdip_DrawPie(pGraphics, pPen, x, y, w, h, StartAngle, SweepAngle)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipDrawPie", Ptr, pGraphics, Ptr, pPen, "float", x, "float", y, "float", w, "float", h, "float", StartAngle, "float", SweepAngle)
}

;#####################################################################################

; Function				Gdip_DrawLine
; Description			This function uses a pen to draw a line into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; x1					x-coordinate of the start of the line
; y1					y-coordinate of the start of the line
; x2					x-coordinate of the end of the line
; y2					y-coordinate of the end of the line
;
; return				status enumeration. 0 = success		

Gdip_DrawLine(pGraphics, pPen, x1, y1, x2, y2)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipDrawLine"
					, Ptr, pGraphics
					, Ptr, pPen
					, "float", x1
					, "float", y1
					, "float", x2
					, "float", y2)
}

;#####################################################################################

; Function				Gdip_DrawLines
; Description			This function uses a pen to draw a series of joined lines into the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pPen					Pointer to a pen
; Points				the coordinates of all the points passed as x1,y1|x2,y2|x3,y3.....
;
; return				status enumeration. 0 = success				

Gdip_DrawLines(pGraphics, pPen, Points)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	StringSplit, Points, Points, |
	VarSetCapacity(PointF, 8*Points0)   
	Loop, %Points0%
	{
		StringSplit, Coord, Points%A_Index%, `,
		NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
	}
	return DllCall("gdiplus\GdipDrawLines", Ptr, pGraphics, Ptr, pPen, Ptr, &PointF, "int", Points0)
}

;#####################################################################################

; Function				Gdip_FillRectangle
; Description			This function uses a brush to fill a rectangle in the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBrush				Pointer to a brush
; x						x-coordinate of the top left of the rectangle
; y						y-coordinate of the top left of the rectangle
; w						width of the rectanlge
; h						height of the rectangle
;
; return				status enumeration. 0 = success

Gdip_FillRectangle(pGraphics, pBrush, x, y, w, h)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipFillRectangle"
					, Ptr, pGraphics
					, Ptr, pBrush
					, "float", x
					, "float", y
					, "float", w
					, "float", h)
}

;#####################################################################################

; Function				Gdip_FillRoundedRectangle
; Description			This function uses a brush to fill a rounded rectangle in the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBrush				Pointer to a brush
; x						x-coordinate of the top left of the rounded rectangle
; y						y-coordinate of the top left of the rounded rectangle
; w						width of the rectanlge
; h						height of the rectangle
; r						radius of the rounded corners
;
; return				status enumeration. 0 = success

Gdip_FillRoundedRectangle(pGraphics, pBrush, x, y, w, h, r)
{
	Region := Gdip_GetClipRegion(pGraphics)
	Gdip_SetClipRect(pGraphics, x-r, y-r, 2*r, 2*r, 4)
	Gdip_SetClipRect(pGraphics, x+w-r, y-r, 2*r, 2*r, 4)
	Gdip_SetClipRect(pGraphics, x-r, y+h-r, 2*r, 2*r, 4)
	Gdip_SetClipRect(pGraphics, x+w-r, y+h-r, 2*r, 2*r, 4)
	E := Gdip_FillRectangle(pGraphics, pBrush, x, y, w, h)
	Gdip_SetClipRegion(pGraphics, Region, 0)
	Gdip_SetClipRect(pGraphics, x-(2*r), y+r, w+(4*r), h-(2*r), 4)
	Gdip_SetClipRect(pGraphics, x+r, y-(2*r), w-(2*r), h+(4*r), 4)
	Gdip_FillEllipse(pGraphics, pBrush, x, y, 2*r, 2*r)
	Gdip_FillEllipse(pGraphics, pBrush, x+w-(2*r), y, 2*r, 2*r)
	Gdip_FillEllipse(pGraphics, pBrush, x, y+h-(2*r), 2*r, 2*r)
	Gdip_FillEllipse(pGraphics, pBrush, x+w-(2*r), y+h-(2*r), 2*r, 2*r)
	Gdip_SetClipRegion(pGraphics, Region, 0)
	Gdip_DeleteRegion(Region)
	return E
}

;#####################################################################################

; Function				Gdip_FillPolygon
; Description			This function uses a brush to fill a polygon in the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBrush				Pointer to a brush
; Points				the coordinates of all the points passed as x1,y1|x2,y2|x3,y3.....
;
; return				status enumeration. 0 = success
;
; notes					Alternate will fill the polygon as a whole, wheras winding will fill each new "segment"
; Alternate 			= 0
; Winding 				= 1

Gdip_FillPolygon(pGraphics, pBrush, Points, FillMode=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	StringSplit, Points, Points, |
	VarSetCapacity(PointF, 8*Points0)   
	Loop, %Points0%
	{
		StringSplit, Coord, Points%A_Index%, `,
		NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
	}   
	return DllCall("gdiplus\GdipFillPolygon", Ptr, pGraphics, Ptr, pBrush, Ptr, &PointF, "int", Points0, "int", FillMode)
}

;#####################################################################################

; Function				Gdip_FillPie
; Description			This function uses a brush to fill a pie in the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBrush				Pointer to a brush
; x						x-coordinate of the top left of the pie
; y						y-coordinate of the top left of the pie
; w						width of the pie
; h						height of the pie
; StartAngle			specifies the angle between the x-axis and the starting point of the pie
; SweepAngle			specifies the angle between the starting and ending points of the pie
;
; return				status enumeration. 0 = success

Gdip_FillPie(pGraphics, pBrush, x, y, w, h, StartAngle, SweepAngle)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipFillPie"
					, Ptr, pGraphics
					, Ptr, pBrush
					, "float", x
					, "float", y
					, "float", w
					, "float", h
					, "float", StartAngle
					, "float", SweepAngle)
}

;#####################################################################################

; Function				Gdip_FillEllipse
; Description			This function uses a brush to fill an ellipse in the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBrush				Pointer to a brush
; x						x-coordinate of the top left of the ellipse
; y						y-coordinate of the top left of the ellipse
; w						width of the ellipse
; h						height of the ellipse
;
; return				status enumeration. 0 = success

Gdip_FillEllipse(pGraphics, pBrush, x, y, w, h)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipFillEllipse", Ptr, pGraphics, Ptr, pBrush, "float", x, "float", y, "float", w, "float", h)
}

;#####################################################################################

; Function				Gdip_FillRegion
; Description			This function uses a brush to fill a region in the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBrush				Pointer to a brush
; Region				Pointer to a Region
;
; return				status enumeration. 0 = success
;
; notes					You can create a region Gdip_CreateRegion() and then add to this

Gdip_FillRegion(pGraphics, pBrush, Region)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipFillRegion", Ptr, pGraphics, Ptr, pBrush, Ptr, Region)
}

;#####################################################################################

; Function				Gdip_FillPath
; Description			This function uses a brush to fill a path in the Graphics of a bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBrush				Pointer to a brush
; Region				Pointer to a Path
;
; return				status enumeration. 0 = success

Gdip_FillPath(pGraphics, pBrush, Path)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipFillPath", Ptr, pGraphics, Ptr, pBrush, Ptr, Path)
}

;#####################################################################################

; Function				Gdip_DrawImagePointsRect
; Description			This function draws a bitmap into the Graphics of another bitmap and skews it
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBitmap				Pointer to a bitmap to be drawn
; Points				Points passed as x1,y1|x2,y2|x3,y3 (3 points: top left, top right, bottom left) describing the drawing of the bitmap
; sx					x-coordinate of source upper-left corner
; sy					y-coordinate of source upper-left corner
; sw					width of source rectangle
; sh					height of source rectangle
; Matrix				a matrix used to alter image attributes when drawing
;
; return				status enumeration. 0 = success
;
; notes					if sx,sy,sw,sh are missed then the entire source bitmap will be used
;						Matrix can be omitted to just draw with no alteration to ARGB
;						Matrix may be passed as a digit from 0 - 1 to change just transparency
;						Matrix can be passed as a matrix with any delimiter

Gdip_DrawImagePointsRect(pGraphics, pBitmap, Points, sx="", sy="", sw="", sh="", Matrix=1)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	StringSplit, Points, Points, |
	VarSetCapacity(PointF, 8*Points0)   
	Loop, %Points0%
	{
		StringSplit, Coord, Points%A_Index%, `,
		NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
	}

	if (Matrix&1 = "")
		ImageAttr := Gdip_SetImageAttributesColorMatrix(Matrix)
	else if (Matrix != 1)
		ImageAttr := Gdip_SetImageAttributesColorMatrix("1|0|0|0|0|0|1|0|0|0|0|0|1|0|0|0|0|0|" Matrix "|0|0|0|0|0|1")
		
	if (sx = "" && sy = "" && sw = "" && sh = "")
	{
		sx := 0, sy := 0
		sw := Gdip_GetImageWidth(pBitmap)
		sh := Gdip_GetImageHeight(pBitmap)
	}

	E := DllCall("gdiplus\GdipDrawImagePointsRect"
				, Ptr, pGraphics
				, Ptr, pBitmap
				, Ptr, &PointF
				, "int", Points0
				, "float", sx
				, "float", sy
				, "float", sw
				, "float", sh
				, "int", 2
				, Ptr, ImageAttr
				, Ptr, 0
				, Ptr, 0)
	if ImageAttr
		Gdip_DisposeImageAttributes(ImageAttr)
	return E
}

;#####################################################################################

; Function				Gdip_DrawImage
; Description			This function draws a bitmap into the Graphics of another bitmap
;
; pGraphics				Pointer to the Graphics of a bitmap
; pBitmap				Pointer to a bitmap to be drawn
; dx					x-coord of destination upper-left corner
; dy					y-coord of destination upper-left corner
; dw					width of destination image
; dh					height of destination image
; sx					x-coordinate of source upper-left corner
; sy					y-coordinate of source upper-left corner
; sw					width of source image
; sh					height of source image
; Matrix				a matrix used to alter image attributes when drawing
;
; return				status enumeration. 0 = success
;
; notes					if sx,sy,sw,sh are missed then the entire source bitmap will be used
;						Gdip_DrawImage performs faster
;						Matrix can be omitted to just draw with no alteration to ARGB
;						Matrix may be passed as a digit from 0 - 1 to change just transparency
;						Matrix can be passed as a matrix with any delimiter. For example:
;						MatrixBright=
;						(
;						1.5		|0		|0		|0		|0
;						0		|1.5	|0		|0		|0
;						0		|0		|1.5	|0		|0
;						0		|0		|0		|1		|0
;						0.05	|0.05	|0.05	|0		|1
;						)
;
; notes					MatrixBright = 1.5|0|0|0|0|0|1.5|0|0|0|0|0|1.5|0|0|0|0|0|1|0|0.05|0.05|0.05|0|1
;						MatrixGreyScale = 0.299|0.299|0.299|0|0|0.587|0.587|0.587|0|0|0.114|0.114|0.114|0|0|0|0|0|1|0|0|0|0|0|1
;						MatrixNegative = -1|0|0|0|0|0|-1|0|0|0|0|0|-1|0|0|0|0|0|1|0|0|0|0|0|1

Gdip_DrawImage(pGraphics, pBitmap, dx="", dy="", dw="", dh="", sx="", sy="", sw="", sh="", Matrix=1)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	if (Matrix&1 = "")
		ImageAttr := Gdip_SetImageAttributesColorMatrix(Matrix)
	else if (Matrix != 1)
		ImageAttr := Gdip_SetImageAttributesColorMatrix("1|0|0|0|0|0|1|0|0|0|0|0|1|0|0|0|0|0|" Matrix "|0|0|0|0|0|1")

	if (sx = "" && sy = "" && sw = "" && sh = "")
	{
		if (dx = "" && dy = "" && dw = "" && dh = "")
		{
			sx := dx := 0, sy := dy := 0
			sw := dw := Gdip_GetImageWidth(pBitmap)
			sh := dh := Gdip_GetImageHeight(pBitmap)
		}
		else
		{
			sx := sy := 0
			sw := Gdip_GetImageWidth(pBitmap)
			sh := Gdip_GetImageHeight(pBitmap)
		}
	}

	E := DllCall("gdiplus\GdipDrawImageRectRect"
				, Ptr, pGraphics
				, Ptr, pBitmap
				, "float", dx
				, "float", dy
				, "float", dw
				, "float", dh
				, "float", sx
				, "float", sy
				, "float", sw
				, "float", sh
				, "int", 2
				, Ptr, ImageAttr
				, Ptr, 0
				, Ptr, 0)
	if ImageAttr
		Gdip_DisposeImageAttributes(ImageAttr)
	return E
}

;#####################################################################################

; Function				Gdip_SetImageAttributesColorMatrix
; Description			This function creates an image matrix ready for drawing
;
; Matrix				a matrix used to alter image attributes when drawing
;						passed with any delimeter
;
; return				returns an image matrix on sucess or 0 if it fails
;
; notes					MatrixBright = 1.5|0|0|0|0|0|1.5|0|0|0|0|0|1.5|0|0|0|0|0|1|0|0.05|0.05|0.05|0|1
;						MatrixGreyScale = 0.299|0.299|0.299|0|0|0.587|0.587|0.587|0|0|0.114|0.114|0.114|0|0|0|0|0|1|0|0|0|0|0|1
;						MatrixNegative = -1|0|0|0|0|0|-1|0|0|0|0|0|-1|0|0|0|0|0|1|0|0|0|0|0|1

Gdip_SetImageAttributesColorMatrix(Matrix)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	VarSetCapacity(ColourMatrix, 100, 0)
	Matrix := RegExReplace(RegExReplace(Matrix, "^[^\d-\.]+([\d\.])", "$1", "", 1), "[^\d-\.]+", "|")
	StringSplit, Matrix, Matrix, |
	Loop, 25
	{
		Matrix := (Matrix%A_Index% != "") ? Matrix%A_Index% : Mod(A_Index-1, 6) ? 0 : 1
		NumPut(Matrix, ColourMatrix, (A_Index-1)*4, "float")
	}
	DllCall("gdiplus\GdipCreateImageAttributes", A_PtrSize ? "UPtr*" : "uint*", ImageAttr)
	DllCall("gdiplus\GdipSetImageAttributesColorMatrix", Ptr, ImageAttr, "int", 1, "int", 1, Ptr, &ColourMatrix, Ptr, 0, "int", 0)
	return ImageAttr
}

;#####################################################################################

; Function				Gdip_GraphicsFromImage
; Description			This function gets the graphics for a bitmap used for drawing functions
;
; pBitmap				Pointer to a bitmap to get the pointer to its graphics
;
; return				returns a pointer to the graphics of a bitmap
;
; notes					a bitmap can be drawn into the graphics of another bitmap

Gdip_GraphicsFromImage(pBitmap)
{
	DllCall("gdiplus\GdipGetImageGraphicsContext", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "UInt*", pGraphics)
	return pGraphics
}

;#####################################################################################

; Function				Gdip_GraphicsFromHDC
; Description			This function gets the graphics from the handle to a device context
;
; hdc					This is the handle to the device context
;
; return				returns a pointer to the graphics of a bitmap
;
; notes					You can draw a bitmap into the graphics of another bitmap

Gdip_GraphicsFromHDC(hdc)
{
    DllCall("gdiplus\GdipCreateFromHDC", A_PtrSize ? "UPtr" : "UInt", hdc, A_PtrSize ? "UPtr*" : "UInt*", pGraphics)
    return pGraphics
}

;#####################################################################################

; Function				Gdip_GetDC
; Description			This function gets the device context of the passed Graphics
;
; hdc					This is the handle to the device context
;
; return				returns the device context for the graphics of a bitmap

Gdip_GetDC(pGraphics)
{
	DllCall("gdiplus\GdipGetDC", A_PtrSize ? "UPtr" : "UInt", pGraphics, A_PtrSize ? "UPtr*" : "UInt*", hdc)
	return hdc
}

;#####################################################################################

; Function				Gdip_ReleaseDC
; Description			This function releases a device context from use for further use
;
; pGraphics				Pointer to the graphics of a bitmap
; hdc					This is the handle to the device context
;
; return				status enumeration. 0 = success

Gdip_ReleaseDC(pGraphics, hdc)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipReleaseDC", Ptr, pGraphics, Ptr, hdc)
}

;#####################################################################################

; Function				Gdip_GraphicsClear
; Description			Clears the graphics of a bitmap ready for further drawing
;
; pGraphics				Pointer to the graphics of a bitmap
; ARGB					The colour to clear the graphics to
;
; return				status enumeration. 0 = success
;
; notes					By default this will make the background invisible
;						Using clipping regions you can clear a particular area on the graphics rather than clearing the entire graphics

Gdip_GraphicsClear(pGraphics, ARGB=0x00ffffff)
{
    return DllCall("gdiplus\GdipGraphicsClear", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", ARGB)
}

;#####################################################################################

; Function				Gdip_BlurBitmap
; Description			Gives a pointer to a blurred bitmap from a pointer to a bitmap
;
; pBitmap				Pointer to a bitmap to be blurred
; Blur					The Amount to blur a bitmap by from 1 (least blur) to 100 (most blur)
;
; return				If the function succeeds, the return value is a pointer to the new blurred bitmap
;						-1 = The blur parameter is outside the range 1-100
;
; notes					This function will not dispose of the original bitmap

Gdip_BlurBitmap(pBitmap, Blur)
{
	if (Blur > 100) || (Blur < 1)
		return -1	
	
	sWidth := Gdip_GetImageWidth(pBitmap), sHeight := Gdip_GetImageHeight(pBitmap)
	dWidth := sWidth//Blur, dHeight := sHeight//Blur

	pBitmap1 := Gdip_CreateBitmap(dWidth, dHeight)
	G1 := Gdip_GraphicsFromImage(pBitmap1)
	Gdip_SetInterpolationMode(G1, 7)
	Gdip_DrawImage(G1, pBitmap, 0, 0, dWidth, dHeight, 0, 0, sWidth, sHeight)

	Gdip_DeleteGraphics(G1)

	pBitmap2 := Gdip_CreateBitmap(sWidth, sHeight)
	G2 := Gdip_GraphicsFromImage(pBitmap2)
	Gdip_SetInterpolationMode(G2, 7)
	Gdip_DrawImage(G2, pBitmap1, 0, 0, sWidth, sHeight, 0, 0, dWidth, dHeight)

	Gdip_DeleteGraphics(G2)
	Gdip_DisposeImage(pBitmap1)
	return pBitmap2
}

;#####################################################################################

; Function:     		Gdip_SaveBitmapToFile
; Description:  		Saves a bitmap to a file in any supported format onto disk
;   
; pBitmap				Pointer to a bitmap
; sOutput      			The name of the file that the bitmap will be saved to. Supported extensions are: .BMP,.DIB,.RLE,.JPG,.JPEG,.JPE,.JFIF,.GIF,.TIF,.TIFF,.PNG
; Quality      			If saving as jpg (.JPG,.JPEG,.JPE,.JFIF) then quality can be 1-100 with default at maximum quality
;
; return      			If the function succeeds, the return value is zero, otherwise:
;						-1 = Extension supplied is not a supported file format
;						-2 = Could not get a list of encoders on system
;						-3 = Could not find matching encoder for specified file format
;						-4 = Could not get WideChar name of output file
;						-5 = Could not save file to disk
;
; notes					This function will use the extension supplied from the sOutput parameter to determine the output format

Gdip_SaveBitmapToFile(pBitmap, sOutput, Quality=75)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	SplitPath, sOutput,,, Extension
	if Extension not in BMP,DIB,RLE,JPG,JPEG,JPE,JFIF,GIF,TIF,TIFF,PNG
		return -1
	Extension := "." Extension

	DllCall("gdiplus\GdipGetImageEncodersSize", "uint*", nCount, "uint*", nSize)
	VarSetCapacity(ci, nSize)
	DllCall("gdiplus\GdipGetImageEncoders", "uint", nCount, "uint", nSize, Ptr, &ci)
	if !(nCount && nSize)
		return -2
	
	If (A_IsUnicode){
		StrGet_Name := "StrGet"
		Loop, %nCount%
		{
			sString := %StrGet_Name%(NumGet(ci, (idx := (48+7*A_PtrSize)*(A_Index-1))+32+3*A_PtrSize), "UTF-16")
			if !InStr(sString, "*" Extension)
				continue
			
			pCodec := &ci+idx
			break
		}
	} else {
		Loop, %nCount%
		{
			Location := NumGet(ci, 76*(A_Index-1)+44)
			nSize := DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "uint", 0, "int",  0, "uint", 0, "uint", 0)
			VarSetCapacity(sString, nSize)
			DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "str", sString, "int", nSize, "uint", 0, "uint", 0)
			if !InStr(sString, "*" Extension)
				continue
			
			pCodec := &ci+76*(A_Index-1)
			break
		}
	}
	
	if !pCodec
		return -3

	if (Quality != 75)
	{
		Quality := (Quality < 0) ? 0 : (Quality > 100) ? 100 : Quality
		if Extension in .JPG,.JPEG,.JPE,.JFIF
		{
			DllCall("gdiplus\GdipGetEncoderParameterListSize", Ptr, pBitmap, Ptr, pCodec, "uint*", nSize)
			VarSetCapacity(EncoderParameters, nSize, 0)
			DllCall("gdiplus\GdipGetEncoderParameterList", Ptr, pBitmap, Ptr, pCodec, "uint", nSize, Ptr, &EncoderParameters)
			Loop, % NumGet(EncoderParameters, "UInt")      ;%
			{
				elem := (24+(A_PtrSize ? A_PtrSize : 4))*(A_Index-1) + 4 + (pad := A_PtrSize = 8 ? 4 : 0)
				if (NumGet(EncoderParameters, elem+16, "UInt") = 1) && (NumGet(EncoderParameters, elem+20, "UInt") = 6)
				{
					p := elem+&EncoderParameters-pad-4
					NumPut(Quality, NumGet(NumPut(4, NumPut(1, p+0)+20, "UInt")), "UInt")
					break
				}
			}      
		}
	}

	if (!A_IsUnicode)
	{
		nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sOutput, "int", -1, Ptr, 0, "int", 0)
		VarSetCapacity(wOutput, nSize*2)
		DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sOutput, "int", -1, Ptr, &wOutput, "int", nSize)
		VarSetCapacity(wOutput, -1)
		if !VarSetCapacity(wOutput)
			return -4
		E := DllCall("gdiplus\GdipSaveImageToFile", Ptr, pBitmap, Ptr, &wOutput, Ptr, pCodec, "uint", p ? p : 0)
	}
	else
		E := DllCall("gdiplus\GdipSaveImageToFile", Ptr, pBitmap, Ptr, &sOutput, Ptr, pCodec, "uint", p ? p : 0)
	return E ? -5 : 0
}

;#####################################################################################

; Function				Gdip_GetPixel
; Description			Gets the ARGB of a pixel in a bitmap
;
; pBitmap				Pointer to a bitmap
; x						x-coordinate of the pixel
; y						y-coordinate of the pixel
;
; return				Returns the ARGB value of the pixel

Gdip_GetPixel(pBitmap, x, y)
{
	DllCall("gdiplus\GdipBitmapGetPixel", A_PtrSize ? "UPtr" : "UInt", pBitmap, "int", x, "int", y, "uint*", ARGB)
	return ARGB
}

;#####################################################################################

; Function				Gdip_SetPixel
; Description			Sets the ARGB of a pixel in a bitmap
;
; pBitmap				Pointer to a bitmap
; x						x-coordinate of the pixel
; y						y-coordinate of the pixel
;
; return				status enumeration. 0 = success

Gdip_SetPixel(pBitmap, x, y, ARGB)
{
   return DllCall("gdiplus\GdipBitmapSetPixel", A_PtrSize ? "UPtr" : "UInt", pBitmap, "int", x, "int", y, "int", ARGB)
}

;#####################################################################################

; Function				Gdip_GetImageWidth
; Description			Gives the width of a bitmap
;
; pBitmap				Pointer to a bitmap
;
; return				Returns the width in pixels of the supplied bitmap

Gdip_GetImageWidth(pBitmap)
{
   DllCall("gdiplus\GdipGetImageWidth", A_PtrSize ? "UPtr" : "UInt", pBitmap, "uint*", Width)
   return Width
}

;#####################################################################################

; Function				Gdip_GetImageHeight
; Description			Gives the height of a bitmap
;
; pBitmap				Pointer to a bitmap
;
; return				Returns the height in pixels of the supplied bitmap

Gdip_GetImageHeight(pBitmap)
{
   DllCall("gdiplus\GdipGetImageHeight", A_PtrSize ? "UPtr" : "UInt", pBitmap, "uint*", Height)
   return Height
}

;#####################################################################################

; Function				Gdip_GetDimensions
; Description			Gives the width and height of a bitmap
;
; pBitmap				Pointer to a bitmap
; Width					ByRef variable. This variable will be set to the width of the bitmap
; Height				ByRef variable. This variable will be set to the height of the bitmap
;
; return				No return value
;						Gdip_GetDimensions(pBitmap, ThisWidth, ThisHeight) will set ThisWidth to the width and ThisHeight to the height

Gdip_GetImageDimensions(pBitmap, ByRef Width, ByRef Height)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	DllCall("gdiplus\GdipGetImageWidth", Ptr, pBitmap, "uint*", Width)
	DllCall("gdiplus\GdipGetImageHeight", Ptr, pBitmap, "uint*", Height)
}

;#####################################################################################

Gdip_GetDimensions(pBitmap, ByRef Width, ByRef Height)
{
	Gdip_GetImageDimensions(pBitmap, Width, Height)
}

;#####################################################################################

Gdip_GetImagePixelFormat(pBitmap)
{
	DllCall("gdiplus\GdipGetImagePixelFormat", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "UInt*", Format)
	return Format
}

;#####################################################################################

; Function				Gdip_GetDpiX
; Description			Gives the horizontal dots per inch of the graphics of a bitmap
;
; pBitmap				Pointer to a bitmap
; Width					ByRef variable. This variable will be set to the width of the bitmap
; Height				ByRef variable. This variable will be set to the height of the bitmap
;
; return				No return value
;						Gdip_GetDimensions(pBitmap, ThisWidth, ThisHeight) will set ThisWidth to the width and ThisHeight to the height

Gdip_GetDpiX(pGraphics)
{
	DllCall("gdiplus\GdipGetDpiX", A_PtrSize ? "UPtr" : "uint", pGraphics, "float*", dpix)
	return Round(dpix)
}

;#####################################################################################

Gdip_GetDpiY(pGraphics)
{
	DllCall("gdiplus\GdipGetDpiY", A_PtrSize ? "UPtr" : "uint", pGraphics, "float*", dpiy)
	return Round(dpiy)
}

;#####################################################################################

Gdip_GetImageHorizontalResolution(pBitmap)
{
	DllCall("gdiplus\GdipGetImageHorizontalResolution", A_PtrSize ? "UPtr" : "uint", pBitmap, "float*", dpix)
	return Round(dpix)
}

;#####################################################################################

Gdip_GetImageVerticalResolution(pBitmap)
{
	DllCall("gdiplus\GdipGetImageVerticalResolution", A_PtrSize ? "UPtr" : "uint", pBitmap, "float*", dpiy)
	return Round(dpiy)
}

;#####################################################################################

Gdip_BitmapSetResolution(pBitmap, dpix, dpiy)
{
	return DllCall("gdiplus\GdipBitmapSetResolution", A_PtrSize ? "UPtr" : "uint", pBitmap, "float", dpix, "float", dpiy)
}

;#####################################################################################

Gdip_CreateBitmapFromFile(sFile, IconNumber=1, IconSize="")
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	, PtrA := A_PtrSize ? "UPtr*" : "UInt*"
	
	SplitPath, sFile,,, ext
	if ext in exe,dll
	{
		Sizes := IconSize ? IconSize : 256 "|" 128 "|" 64 "|" 48 "|" 32 "|" 16
		BufSize := 16 + (2*(A_PtrSize ? A_PtrSize : 4))
		
		VarSetCapacity(buf, BufSize, 0)
		Loop, Parse, Sizes, |
		{
			DllCall("PrivateExtractIcons", "str", sFile, "int", IconNumber-1, "int", A_LoopField, "int", A_LoopField, PtrA, hIcon, PtrA, 0, "uint", 1, "uint", 0)
			
			if !hIcon
				continue

			if !DllCall("GetIconInfo", Ptr, hIcon, Ptr, &buf)
			{
				DestroyIcon(hIcon)
				continue
			}
			
			hbmMask  := NumGet(buf, 12 + ((A_PtrSize ? A_PtrSize : 4) - 4))
			hbmColor := NumGet(buf, 12 + ((A_PtrSize ? A_PtrSize : 4) - 4) + (A_PtrSize ? A_PtrSize : 4))
			if !(hbmColor && DllCall("GetObject", Ptr, hbmColor, "int", BufSize, Ptr, &buf))
			{
				DestroyIcon(hIcon)
				continue
			}
			break
		}
		if !hIcon
			return -1

		Width := NumGet(buf, 4, "int"), Height := NumGet(buf, 8, "int")
		hbm := CreateDIBSection(Width, -Height), hdc := CreateCompatibleDC(), obm := SelectObject(hdc, hbm)
		if !DllCall("DrawIconEx", Ptr, hdc, "int", 0, "int", 0, Ptr, hIcon, "uint", Width, "uint", Height, "uint", 0, Ptr, 0, "uint", 3)
		{
			DestroyIcon(hIcon)
			return -2
		}
		
		VarSetCapacity(dib, 104)
		DllCall("GetObject", Ptr, hbm, "int", A_PtrSize = 8 ? 104 : 84, Ptr, &dib) ; sizeof(DIBSECTION) = 76+2*(A_PtrSize=8?4:0)+2*A_PtrSize
		Stride := NumGet(dib, 12, "Int"), Bits := NumGet(dib, 20 + (A_PtrSize = 8 ? 4 : 0)) ; padding
		DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", Width, "int", Height, "int", Stride, "int", 0x26200A, Ptr, Bits, PtrA, pBitmapOld)
		pBitmap := Gdip_CreateBitmap(Width, Height)
		G := Gdip_GraphicsFromImage(pBitmap)
		, Gdip_DrawImage(G, pBitmapOld, 0, 0, Width, Height, 0, 0, Width, Height)
		SelectObject(hdc, obm), DeleteObject(hbm), DeleteDC(hdc)
		Gdip_DeleteGraphics(G), Gdip_DisposeImage(pBitmapOld)
		DestroyIcon(hIcon)
	}
	else
	{
		if (!A_IsUnicode)
		{
			VarSetCapacity(wFile, 1024)
			DllCall("kernel32\MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sFile, "int", -1, Ptr, &wFile, "int", 512)
			DllCall("gdiplus\GdipCreateBitmapFromFile", Ptr, &wFile, PtrA, pBitmap)
		}
		else
			DllCall("gdiplus\GdipCreateBitmapFromFile", Ptr, &sFile, PtrA, pBitmap)
	}
	
	return pBitmap
}

;#####################################################################################

Gdip_CreateBitmapFromHBITMAP(hBitmap, Palette=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", Ptr, hBitmap, Ptr, Palette, A_PtrSize ? "UPtr*" : "uint*", pBitmap)
	return pBitmap
}

;#####################################################################################

Gdip_CreateHBITMAPFromBitmap(pBitmap, Background=0xffffffff)
{
	DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hbm, "int", Background)
	return hbm
}

;#####################################################################################

Gdip_CreateBitmapFromHICON(hIcon)
{
	DllCall("gdiplus\GdipCreateBitmapFromHICON", A_PtrSize ? "UPtr" : "UInt", hIcon, A_PtrSize ? "UPtr*" : "uint*", pBitmap)
	return pBitmap
}

;#####################################################################################

Gdip_CreateHICONFromBitmap(pBitmap)
{
	DllCall("gdiplus\GdipCreateHICONFromBitmap", A_PtrSize ? "UPtr" : "UInt", pBitmap, A_PtrSize ? "UPtr*" : "uint*", hIcon)
	return hIcon
}

;#####################################################################################

Gdip_CreateBitmap(Width, Height, Format=0x26200A)
{
    DllCall("gdiplus\GdipCreateBitmapFromScan0", "int", Width, "int", Height, "int", 0, "int", Format, A_PtrSize ? "UPtr" : "UInt", 0, A_PtrSize ? "UPtr*" : "uint*", pBitmap)
    Return pBitmap
}

;#####################################################################################

Gdip_CreateBitmapFromClipboard()
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	if !DllCall("OpenClipboard", Ptr, 0)
		return -1
	if !DllCall("IsClipboardFormatAvailable", "uint", 8)
		return -2
	if !hBitmap := DllCall("GetClipboardData", "uint", 2, Ptr)
		return -3
	if !pBitmap := Gdip_CreateBitmapFromHBITMAP(hBitmap)
		return -4
	if !DllCall("CloseClipboard")
		return -5
	DeleteObject(hBitmap)
	return pBitmap
}

;#####################################################################################

Gdip_SetBitmapToClipboard(pBitmap)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	off1 := A_PtrSize = 8 ? 52 : 44, off2 := A_PtrSize = 8 ? 32 : 24
	hBitmap := Gdip_CreateHBITMAPFromBitmap(pBitmap)
	DllCall("GetObject", Ptr, hBitmap, "int", VarSetCapacity(oi, A_PtrSize = 8 ? 104 : 84, 0), Ptr, &oi)
	hdib := DllCall("GlobalAlloc", "uint", 2, Ptr, 40+NumGet(oi, off1, "UInt"), Ptr)
	pdib := DllCall("GlobalLock", Ptr, hdib, Ptr)
	DllCall("RtlMoveMemory", Ptr, pdib, Ptr, &oi+off2, Ptr, 40)
	DllCall("RtlMoveMemory", Ptr, pdib+40, Ptr, NumGet(oi, off2 - (A_PtrSize ? A_PtrSize : 4), Ptr), Ptr, NumGet(oi, off1, "UInt"))
	DllCall("GlobalUnlock", Ptr, hdib)
	DllCall("DeleteObject", Ptr, hBitmap)
	DllCall("OpenClipboard", Ptr, 0)
	DllCall("EmptyClipboard")
	DllCall("SetClipboardData", "uint", 8, Ptr, hdib)
	DllCall("CloseClipboard")
}

;#####################################################################################

Gdip_CloneBitmapArea(pBitmap, x, y, w, h, Format=0x26200A)
{
	DllCall("gdiplus\GdipCloneBitmapArea"
					, "float", x
					, "float", y
					, "float", w
					, "float", h
					, "int", Format
					, A_PtrSize ? "UPtr" : "UInt", pBitmap
					, A_PtrSize ? "UPtr*" : "UInt*", pBitmapDest)
	return pBitmapDest
}

;#####################################################################################
; Create resources
;#####################################################################################

Gdip_CreatePen(ARGB, w)
{
   DllCall("gdiplus\GdipCreatePen1", "UInt", ARGB, "float", w, "int", 2, A_PtrSize ? "UPtr*" : "UInt*", pPen)
   return pPen
}

;#####################################################################################

Gdip_CreatePenFromBrush(pBrush, w)
{
	DllCall("gdiplus\GdipCreatePen2", A_PtrSize ? "UPtr" : "UInt", pBrush, "float", w, "int", 2, A_PtrSize ? "UPtr*" : "UInt*", pPen)
	return pPen
}

;#####################################################################################

Gdip_BrushCreateSolid(ARGB=0xff000000)
{
	DllCall("gdiplus\GdipCreateSolidFill", "UInt", ARGB, A_PtrSize ? "UPtr*" : "UInt*", pBrush)
	return pBrush
}

;#####################################################################################

; HatchStyleHorizontal = 0
; HatchStyleVertical = 1
; HatchStyleForwardDiagonal = 2
; HatchStyleBackwardDiagonal = 3
; HatchStyleCross = 4
; HatchStyleDiagonalCross = 5
; HatchStyle05Percent = 6
; HatchStyle10Percent = 7
; HatchStyle20Percent = 8
; HatchStyle25Percent = 9
; HatchStyle30Percent = 10
; HatchStyle40Percent = 11
; HatchStyle50Percent = 12
; HatchStyle60Percent = 13
; HatchStyle70Percent = 14
; HatchStyle75Percent = 15
; HatchStyle80Percent = 16
; HatchStyle90Percent = 17
; HatchStyleLightDownwardDiagonal = 18
; HatchStyleLightUpwardDiagonal = 19
; HatchStyleDarkDownwardDiagonal = 20
; HatchStyleDarkUpwardDiagonal = 21
; HatchStyleWideDownwardDiagonal = 22
; HatchStyleWideUpwardDiagonal = 23
; HatchStyleLightVertical = 24
; HatchStyleLightHorizontal = 25
; HatchStyleNarrowVertical = 26
; HatchStyleNarrowHorizontal = 27
; HatchStyleDarkVertical = 28
; HatchStyleDarkHorizontal = 29
; HatchStyleDashedDownwardDiagonal = 30
; HatchStyleDashedUpwardDiagonal = 31
; HatchStyleDashedHorizontal = 32
; HatchStyleDashedVertical = 33
; HatchStyleSmallConfetti = 34
; HatchStyleLargeConfetti = 35
; HatchStyleZigZag = 36
; HatchStyleWave = 37
; HatchStyleDiagonalBrick = 38
; HatchStyleHorizontalBrick = 39
; HatchStyleWeave = 40
; HatchStylePlaid = 41
; HatchStyleDivot = 42
; HatchStyleDottedGrid = 43
; HatchStyleDottedDiamond = 44
; HatchStyleShingle = 45
; HatchStyleTrellis = 46
; HatchStyleSphere = 47
; HatchStyleSmallGrid = 48
; HatchStyleSmallCheckerBoard = 49
; HatchStyleLargeCheckerBoard = 50
; HatchStyleOutlinedDiamond = 51
; HatchStyleSolidDiamond = 52
; HatchStyleTotal = 53
Gdip_BrushCreateHatch(ARGBfront, ARGBback, HatchStyle=0)
{
	DllCall("gdiplus\GdipCreateHatchBrush", "int", HatchStyle, "UInt", ARGBfront, "UInt", ARGBback, A_PtrSize ? "UPtr*" : "UInt*", pBrush)
	return pBrush
}

;#####################################################################################

Gdip_CreateTextureBrush(pBitmap, WrapMode=1, x=0, y=0, w="", h="")
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	, PtrA := A_PtrSize ? "UPtr*" : "UInt*"
	
	if !(w && h)
		DllCall("gdiplus\GdipCreateTexture", Ptr, pBitmap, "int", WrapMode, PtrA, pBrush)
	else
		DllCall("gdiplus\GdipCreateTexture2", Ptr, pBitmap, "int", WrapMode, "float", x, "float", y, "float", w, "float", h, PtrA, pBrush)
	return pBrush
}

;#####################################################################################

; WrapModeTile = 0
; WrapModeTileFlipX = 1
; WrapModeTileFlipY = 2
; WrapModeTileFlipXY = 3
; WrapModeClamp = 4
Gdip_CreateLineBrush(x1, y1, x2, y2, ARGB1, ARGB2, WrapMode=1)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	CreatePointF(PointF1, x1, y1), CreatePointF(PointF2, x2, y2)
	DllCall("gdiplus\GdipCreateLineBrush", Ptr, &PointF1, Ptr, &PointF2, "Uint", ARGB1, "Uint", ARGB2, "int", WrapMode, A_PtrSize ? "UPtr*" : "UInt*", LGpBrush)
	return LGpBrush
}

;#####################################################################################

; LinearGradientModeHorizontal = 0
; LinearGradientModeVertical = 1
; LinearGradientModeForwardDiagonal = 2
; LinearGradientModeBackwardDiagonal = 3
Gdip_CreateLineBrushFromRect(x, y, w, h, ARGB1, ARGB2, LinearGradientMode=1, WrapMode=1)
{
	CreateRectF(RectF, x, y, w, h)
	DllCall("gdiplus\GdipCreateLineBrushFromRect", A_PtrSize ? "UPtr" : "UInt", &RectF, "int", ARGB1, "int", ARGB2, "int", LinearGradientMode, "int", WrapMode, A_PtrSize ? "UPtr*" : "UInt*", LGpBrush)
	return LGpBrush
}

;#####################################################################################

Gdip_CloneBrush(pBrush)
{
	DllCall("gdiplus\GdipCloneBrush", A_PtrSize ? "UPtr" : "UInt", pBrush, A_PtrSize ? "UPtr*" : "UInt*", pBrushClone)
	return pBrushClone
}

;#####################################################################################
; Delete resources
;#####################################################################################

Gdip_DeletePen(pPen)
{
   return DllCall("gdiplus\GdipDeletePen", A_PtrSize ? "UPtr" : "UInt", pPen)
}

;#####################################################################################

Gdip_DeleteBrush(pBrush)
{
   return DllCall("gdiplus\GdipDeleteBrush", A_PtrSize ? "UPtr" : "UInt", pBrush)
}

;#####################################################################################

Gdip_DisposeImage(pBitmap)
{
   return DllCall("gdiplus\GdipDisposeImage", A_PtrSize ? "UPtr" : "UInt", pBitmap)
}

;#####################################################################################

Gdip_DeleteGraphics(pGraphics)
{
   return DllCall("gdiplus\GdipDeleteGraphics", A_PtrSize ? "UPtr" : "UInt", pGraphics)
}

;#####################################################################################

Gdip_DisposeImageAttributes(ImageAttr)
{
	return DllCall("gdiplus\GdipDisposeImageAttributes", A_PtrSize ? "UPtr" : "UInt", ImageAttr)
}

;#####################################################################################

Gdip_DeleteFont(hFont)
{
   return DllCall("gdiplus\GdipDeleteFont", A_PtrSize ? "UPtr" : "UInt", hFont)
}

;#####################################################################################

Gdip_DeleteStringFormat(hFormat)
{
   return DllCall("gdiplus\GdipDeleteStringFormat", A_PtrSize ? "UPtr" : "UInt", hFormat)
}

;#####################################################################################

Gdip_DeleteFontFamily(hFamily)
{
   return DllCall("gdiplus\GdipDeleteFontFamily", A_PtrSize ? "UPtr" : "UInt", hFamily)
}

;#####################################################################################

Gdip_DeleteMatrix(Matrix)
{
   return DllCall("gdiplus\GdipDeleteMatrix", A_PtrSize ? "UPtr" : "UInt", Matrix)
}

;#####################################################################################
; Text functions
;#####################################################################################

Gdip_TextToGraphics(pGraphics, Text, Options, Font="Arial", Width="", Height="", Measure=0)
{
	IWidth := Width, IHeight:= Height
	
	RegExMatch(Options, "i)X([\-\d\.]+)(p*)", xpos)
	RegExMatch(Options, "i)Y([\-\d\.]+)(p*)", ypos)
	RegExMatch(Options, "i)W([\-\d\.]+)(p*)", Width)
	RegExMatch(Options, "i)H([\-\d\.]+)(p*)", Height)
	RegExMatch(Options, "i)C(?!(entre|enter))([a-f\d]+)", Colour)
	RegExMatch(Options, "i)Top|Up|Bottom|Down|vCentre|vCenter", vPos)
	RegExMatch(Options, "i)NoWrap", NoWrap)
	RegExMatch(Options, "i)R(\d)", Rendering)
	RegExMatch(Options, "i)S(\d+)(p*)", Size)

	if !Gdip_DeleteBrush(Gdip_CloneBrush(Colour2))
		PassBrush := 1, pBrush := Colour2
	
	if !(IWidth && IHeight) && (xpos2 || ypos2 || Width2 || Height2 || Size2)
		return -1

	Style := 0, Styles := "Regular|Bold|Italic|BoldItalic|Underline|Strikeout"
	Loop, Parse, Styles, |
	{
		if RegExMatch(Options, "\b" A_loopField)
		Style |= (A_LoopField != "StrikeOut") ? (A_Index-1) : 8
	}
  
	Align := 0, Alignments := "Near|Left|Centre|Center|Far|Right"
	Loop, Parse, Alignments, |
	{
		if RegExMatch(Options, "\b" A_loopField)
			Align |= A_Index//2.1      ; 0|0|1|1|2|2
	}

	xpos := (xpos1 != "") ? xpos2 ? IWidth*(xpos1/100) : xpos1 : 0
	ypos := (ypos1 != "") ? ypos2 ? IHeight*(ypos1/100) : ypos1 : 0
	Width := Width1 ? Width2 ? IWidth*(Width1/100) : Width1 : IWidth
	Height := Height1 ? Height2 ? IHeight*(Height1/100) : Height1 : IHeight
	if !PassBrush
		Colour := "0x" (Colour2 ? Colour2 : "ff000000")
	Rendering := ((Rendering1 >= 0) && (Rendering1 <= 5)) ? Rendering1 : 4
	Size := (Size1 > 0) ? Size2 ? IHeight*(Size1/100) : Size1 : 12

	hFamily := Gdip_FontFamilyCreate(Font)
	hFont := Gdip_FontCreate(hFamily, Size, Style)
	FormatStyle := NoWrap ? 0x4000 | 0x1000 : 0x4000
	hFormat := Gdip_StringFormatCreate(FormatStyle)
	pBrush := PassBrush ? pBrush : Gdip_BrushCreateSolid(Colour)
	if !(hFamily && hFont && hFormat && pBrush && pGraphics)
		return !pGraphics ? -2 : !hFamily ? -3 : !hFont ? -4 : !hFormat ? -5 : !pBrush ? -6 : 0
   
	CreateRectF(RC, xpos, ypos, Width, Height)
	Gdip_SetStringFormatAlign(hFormat, Align)
	Gdip_SetTextRenderingHint(pGraphics, Rendering)
	ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, RC)

	if vPos
	{
		StringSplit, ReturnRC, ReturnRC, |
		
		if (vPos = "vCentre") || (vPos = "vCenter")
			ypos += (Height-ReturnRC4)//2
		else if (vPos = "Top") || (vPos = "Up")
			ypos := 0
		else if (vPos = "Bottom") || (vPos = "Down")
			ypos := Height-ReturnRC4
		
		CreateRectF(RC, xpos, ypos, Width, ReturnRC4)
		ReturnRC := Gdip_MeasureString(pGraphics, Text, hFont, hFormat, RC)
	}

	if !Measure
		E := Gdip_DrawString(pGraphics, Text, hFont, hFormat, pBrush, RC)

	if !PassBrush
		Gdip_DeleteBrush(pBrush)
	Gdip_DeleteStringFormat(hFormat)   
	Gdip_DeleteFont(hFont)
	Gdip_DeleteFontFamily(hFamily)
	return E ? E : ReturnRC
}

;#####################################################################################

Gdip_DrawString(pGraphics, sString, hFont, hFormat, pBrush, ByRef RectF)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	if (!A_IsUnicode)
	{
		nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, Ptr, 0, "int", 0)
		VarSetCapacity(wString, nSize*2)
		DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, Ptr, &wString, "int", nSize)
	}
	
	return DllCall("gdiplus\GdipDrawString"
					, Ptr, pGraphics
					, Ptr, A_IsUnicode ? &sString : &wString
					, "int", -1
					, Ptr, hFont
					, Ptr, &RectF
					, Ptr, hFormat
					, Ptr, pBrush)
}

;#####################################################################################

Gdip_MeasureString(pGraphics, sString, hFont, hFormat, ByRef RectF)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	VarSetCapacity(RC, 16)
	if !A_IsUnicode
	{
		nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, "uint", 0, "int", 0)
		VarSetCapacity(wString, nSize*2)   
		DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &sString, "int", -1, Ptr, &wString, "int", nSize)
	}
	
	DllCall("gdiplus\GdipMeasureString"
					, Ptr, pGraphics
					, Ptr, A_IsUnicode ? &sString : &wString
					, "int", -1
					, Ptr, hFont
					, Ptr, &RectF
					, Ptr, hFormat
					, Ptr, &RC
					, "uint*", Chars
					, "uint*", Lines)
	
	return &RC ? NumGet(RC, 0, "float") "|" NumGet(RC, 4, "float") "|" NumGet(RC, 8, "float") "|" NumGet(RC, 12, "float") "|" Chars "|" Lines : 0
}

; Near = 0
; Center = 1
; Far = 2
Gdip_SetStringFormatAlign(hFormat, Align)
{
   return DllCall("gdiplus\GdipSetStringFormatAlign", A_PtrSize ? "UPtr" : "UInt", hFormat, "int", Align)
}

; StringFormatFlagsDirectionRightToLeft    = 0x00000001
; StringFormatFlagsDirectionVertical       = 0x00000002
; StringFormatFlagsNoFitBlackBox           = 0x00000004
; StringFormatFlagsDisplayFormatControl    = 0x00000020
; StringFormatFlagsNoFontFallback          = 0x00000400
; StringFormatFlagsMeasureTrailingSpaces   = 0x00000800
; StringFormatFlagsNoWrap                  = 0x00001000
; StringFormatFlagsLineLimit               = 0x00002000
; StringFormatFlagsNoClip                  = 0x00004000 
Gdip_StringFormatCreate(Format=0, Lang=0)
{
   DllCall("gdiplus\GdipCreateStringFormat", "int", Format, "int", Lang, A_PtrSize ? "UPtr*" : "UInt*", hFormat)
   return hFormat
}

; Regular = 0
; Bold = 1
; Italic = 2
; BoldItalic = 3
; Underline = 4
; Strikeout = 8
Gdip_FontCreate(hFamily, Size, Style=0)
{
   DllCall("gdiplus\GdipCreateFont", A_PtrSize ? "UPtr" : "UInt", hFamily, "float", Size, "int", Style, "int", 0, A_PtrSize ? "UPtr*" : "UInt*", hFont)
   return hFont
}

Gdip_FontFamilyCreate(Font)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	if (!A_IsUnicode)
	{
		nSize := DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &Font, "int", -1, "uint", 0, "int", 0)
		VarSetCapacity(wFont, nSize*2)
		DllCall("MultiByteToWideChar", "uint", 0, "uint", 0, Ptr, &Font, "int", -1, Ptr, &wFont, "int", nSize)
	}
	
	DllCall("gdiplus\GdipCreateFontFamilyFromName"
					, Ptr, A_IsUnicode ? &Font : &wFont
					, "uint", 0
					, A_PtrSize ? "UPtr*" : "UInt*", hFamily)
	
	return hFamily
}

;#####################################################################################
; Matrix functions
;#####################################################################################

Gdip_CreateAffineMatrix(m11, m12, m21, m22, x, y)
{
   DllCall("gdiplus\GdipCreateMatrix2", "float", m11, "float", m12, "float", m21, "float", m22, "float", x, "float", y, A_PtrSize ? "UPtr*" : "UInt*", Matrix)
   return Matrix
}

Gdip_CreateMatrix()
{
   DllCall("gdiplus\GdipCreateMatrix", A_PtrSize ? "UPtr*" : "UInt*", Matrix)
   return Matrix
}

;#####################################################################################
; GraphicsPath functions
;#####################################################################################

; Alternate = 0
; Winding = 1
Gdip_CreatePath(BrushMode=0)
{
	DllCall("gdiplus\GdipCreatePath", "int", BrushMode, A_PtrSize ? "UPtr*" : "UInt*", Path)
	return Path
}

Gdip_AddPathEllipse(Path, x, y, w, h)
{
	return DllCall("gdiplus\GdipAddPathEllipse", A_PtrSize ? "UPtr" : "UInt", Path, "float", x, "float", y, "float", w, "float", h)
}

Gdip_AddPathPolygon(Path, Points)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	StringSplit, Points, Points, |
	VarSetCapacity(PointF, 8*Points0)   
	Loop, %Points0%
	{
		StringSplit, Coord, Points%A_Index%, `,
		NumPut(Coord1, PointF, 8*(A_Index-1), "float"), NumPut(Coord2, PointF, (8*(A_Index-1))+4, "float")
	}   

	return DllCall("gdiplus\GdipAddPathPolygon", Ptr, Path, Ptr, &PointF, "int", Points0)
}

Gdip_DeletePath(Path)
{
	return DllCall("gdiplus\GdipDeletePath", A_PtrSize ? "UPtr" : "UInt", Path)
}

;#####################################################################################
; Quality functions
;#####################################################################################

; SystemDefault = 0
; SingleBitPerPixelGridFit = 1
; SingleBitPerPixel = 2
; AntiAliasGridFit = 3
; AntiAlias = 4
Gdip_SetTextRenderingHint(pGraphics, RenderingHint)
{
	return DllCall("gdiplus\GdipSetTextRenderingHint", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", RenderingHint)
}

; Default = 0
; LowQuality = 1
; HighQuality = 2
; Bilinear = 3
; Bicubic = 4
; NearestNeighbor = 5
; HighQualityBilinear = 6
; HighQualityBicubic = 7
Gdip_SetInterpolationMode(pGraphics, InterpolationMode)
{
   return DllCall("gdiplus\GdipSetInterpolationMode", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", InterpolationMode)
}

; Default = 0
; HighSpeed = 1
; HighQuality = 2
; None = 3
; AntiAlias = 4
Gdip_SetSmoothingMode(pGraphics, SmoothingMode)
{
   return DllCall("gdiplus\GdipSetSmoothingMode", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", SmoothingMode)
}

; CompositingModeSourceOver = 0 (blended)
; CompositingModeSourceCopy = 1 (overwrite)
Gdip_SetCompositingMode(pGraphics, CompositingMode=0)
{
   return DllCall("gdiplus\GdipSetCompositingMode", A_PtrSize ? "UPtr" : "UInt", pGraphics, "int", CompositingMode)
}

;#####################################################################################
; Extra functions
;#####################################################################################

Gdip_Startup()
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	if !DllCall("GetModuleHandle", "str", "gdiplus", Ptr)
		DllCall("LoadLibrary", "str", "gdiplus")
	VarSetCapacity(si, A_PtrSize = 8 ? 24 : 16, 0), si := Chr(1)
	DllCall("gdiplus\GdiplusStartup", A_PtrSize ? "UPtr*" : "uint*", pToken, Ptr, &si, Ptr, 0)
	return pToken
}

Gdip_Shutdown(pToken)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	DllCall("gdiplus\GdiplusShutdown", Ptr, pToken)
	if hModule := DllCall("GetModuleHandle", "str", "gdiplus", Ptr)
		DllCall("FreeLibrary", Ptr, hModule)
	return 0
}

; Prepend = 0; The new operation is applied before the old operation.
; Append = 1; The new operation is applied after the old operation.
Gdip_RotateWorldTransform(pGraphics, Angle, MatrixOrder=0)
{
	return DllCall("gdiplus\GdipRotateWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", Angle, "int", MatrixOrder)
}

Gdip_ScaleWorldTransform(pGraphics, x, y, MatrixOrder=0)
{
	return DllCall("gdiplus\GdipScaleWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", x, "float", y, "int", MatrixOrder)
}

Gdip_TranslateWorldTransform(pGraphics, x, y, MatrixOrder=0)
{
	return DllCall("gdiplus\GdipTranslateWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", x, "float", y, "int", MatrixOrder)
}

Gdip_ResetWorldTransform(pGraphics)
{
	return DllCall("gdiplus\GdipResetWorldTransform", A_PtrSize ? "UPtr" : "UInt", pGraphics)
}

Gdip_GetRotatedTranslation(Width, Height, Angle, ByRef xTranslation, ByRef yTranslation)
{
	pi := 3.14159, TAngle := Angle*(pi/180)	

	Bound := (Angle >= 0) ? Mod(Angle, 360) : 360-Mod(-Angle, -360)
	if ((Bound >= 0) && (Bound <= 90))
		xTranslation := Height*Sin(TAngle), yTranslation := 0
	else if ((Bound > 90) && (Bound <= 180))
		xTranslation := (Height*Sin(TAngle))-(Width*Cos(TAngle)), yTranslation := -Height*Cos(TAngle)
	else if ((Bound > 180) && (Bound <= 270))
		xTranslation := -(Width*Cos(TAngle)), yTranslation := -(Height*Cos(TAngle))-(Width*Sin(TAngle))
	else if ((Bound > 270) && (Bound <= 360))
		xTranslation := 0, yTranslation := -Width*Sin(TAngle)
}

Gdip_GetRotatedDimensions(Width, Height, Angle, ByRef RWidth, ByRef RHeight)
{
	pi := 3.14159, TAngle := Angle*(pi/180)
	if !(Width && Height)
		return -1
	RWidth := Ceil(Abs(Width*Cos(TAngle))+Abs(Height*Sin(TAngle)))
	RHeight := Ceil(Abs(Width*Sin(TAngle))+Abs(Height*Cos(Tangle)))
}

; RotateNoneFlipNone   = 0
; Rotate90FlipNone     = 1
; Rotate180FlipNone    = 2
; Rotate270FlipNone    = 3
; RotateNoneFlipX      = 4
; Rotate90FlipX        = 5
; Rotate180FlipX       = 6
; Rotate270FlipX       = 7
; RotateNoneFlipY      = Rotate180FlipX
; Rotate90FlipY        = Rotate270FlipX
; Rotate180FlipY       = RotateNoneFlipX
; Rotate270FlipY       = Rotate90FlipX
; RotateNoneFlipXY     = Rotate180FlipNone
; Rotate90FlipXY       = Rotate270FlipNone
; Rotate180FlipXY      = RotateNoneFlipNone
; Rotate270FlipXY      = Rotate90FlipNone 

Gdip_ImageRotateFlip(pBitmap, RotateFlipType=1)
{
	return DllCall("gdiplus\GdipImageRotateFlip", A_PtrSize ? "UPtr" : "UInt", pBitmap, "int", RotateFlipType)
}

; Replace = 0
; Intersect = 1
; Union = 2
; Xor = 3
; Exclude = 4
; Complement = 5
Gdip_SetClipRect(pGraphics, x, y, w, h, CombineMode=0)
{
   return DllCall("gdiplus\GdipSetClipRect",  A_PtrSize ? "UPtr" : "UInt", pGraphics, "float", x, "float", y, "float", w, "float", h, "int", CombineMode)
}

Gdip_SetClipPath(pGraphics, Path, CombineMode=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	return DllCall("gdiplus\GdipSetClipPath", Ptr, pGraphics, Ptr, Path, "int", CombineMode)
}

Gdip_ResetClip(pGraphics)
{
   return DllCall("gdiplus\GdipResetClip", A_PtrSize ? "UPtr" : "UInt", pGraphics)
}

Gdip_GetClipRegion(pGraphics)
{
	Region := Gdip_CreateRegion()
	DllCall("gdiplus\GdipGetClip", A_PtrSize ? "UPtr" : "UInt", pGraphics, "UInt*", Region)
	return Region
}

Gdip_SetClipRegion(pGraphics, Region, CombineMode=0)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("gdiplus\GdipSetClipRegion", Ptr, pGraphics, Ptr, Region, "int", CombineMode)
}

Gdip_CreateRegion()
{
	DllCall("gdiplus\GdipCreateRegion", "UInt*", Region)
	return Region
}

Gdip_DeleteRegion(Region)
{
	return DllCall("gdiplus\GdipDeleteRegion", A_PtrSize ? "UPtr" : "UInt", Region)
}

;#####################################################################################
; BitmapLockBits
;#####################################################################################

Gdip_LockBits(pBitmap, x, y, w, h, ByRef Stride, ByRef Scan0, ByRef BitmapData, LockMode = 3, PixelFormat = 0x26200a)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	CreateRect(Rect, x, y, w, h)
	VarSetCapacity(BitmapData, 16+2*(A_PtrSize ? A_PtrSize : 4), 0)
	E := DllCall("Gdiplus\GdipBitmapLockBits", Ptr, pBitmap, Ptr, &Rect, "uint", LockMode, "int", PixelFormat, Ptr, &BitmapData)
	Stride := NumGet(BitmapData, 8, "Int")
	Scan0 := NumGet(BitmapData, 16, Ptr)
	return E
}

;#####################################################################################

Gdip_UnlockBits(pBitmap, ByRef BitmapData)
{
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	return DllCall("Gdiplus\GdipBitmapUnlockBits", Ptr, pBitmap, Ptr, &BitmapData)
}

;#####################################################################################

Gdip_SetLockBitPixel(ARGB, Scan0, x, y, Stride)
{
	Numput(ARGB, Scan0+0, (x*4)+(y*Stride), "UInt")
}

;#####################################################################################

Gdip_GetLockBitPixel(Scan0, x, y, Stride)
{
	return NumGet(Scan0+0, (x*4)+(y*Stride), "UInt")
}

;#####################################################################################

Gdip_PixelateBitmap(pBitmap, ByRef pBitmapOut, BlockSize)
{
	static PixelateBitmap
	
	Ptr := A_PtrSize ? "UPtr" : "UInt"
	
	if (!PixelateBitmap)
	{
		if A_PtrSize != 8 ; x86 machine code
		MCode_PixelateBitmap =
		(LTrim Join
		558BEC83EC3C8B4514538B5D1C99F7FB56578BC88955EC894DD885C90F8E830200008B451099F7FB8365DC008365E000894DC88955F08945E833FF897DD4
		397DE80F8E160100008BCB0FAFCB894DCC33C08945F88945FC89451C8945143BD87E608B45088D50028BC82BCA8BF02BF2418945F48B45E02955F4894DC4
		8D0CB80FAFCB03CA895DD08BD1895DE40FB64416030145140FB60201451C8B45C40FB604100145FC8B45F40FB604020145F883C204FF4DE475D6034D18FF
		4DD075C98B4DCC8B451499F7F98945148B451C99F7F989451C8B45FC99F7F98945FC8B45F899F7F98945F885DB7E648B450C8D50028BC82BCA83C103894D
		C48BC82BCA41894DF48B4DD48945E48B45E02955E48D0C880FAFCB03CA895DD08BD18BF38A45148B7DC48804178A451C8B7DF488028A45FC8804178A45F8
		8B7DE488043A83C2044E75DA034D18FF4DD075CE8B4DCC8B7DD447897DD43B7DE80F8CF2FEFFFF837DF0000F842C01000033C08945F88945FC89451C8945
		148945E43BD87E65837DF0007E578B4DDC034DE48B75E80FAF4D180FAFF38B45088D500203CA8D0CB18BF08BF88945F48B45F02BF22BFA2955F48945CC0F
		B6440E030145140FB60101451C0FB6440F010145FC8B45F40FB604010145F883C104FF4DCC75D8FF45E4395DE47C9B8B4DF00FAFCB85C9740B8B451499F7
		F9894514EB048365140033F63BCE740B8B451C99F7F989451CEB0389751C3BCE740B8B45FC99F7F98945FCEB038975FC3BCE740B8B45F899F7F98945F8EB
		038975F88975E43BDE7E5A837DF0007E4C8B4DDC034DE48B75E80FAF4D180FAFF38B450C8D500203CA8D0CB18BF08BF82BF22BFA2BC28B55F08955CC8A55
		1488540E038A551C88118A55FC88540F018A55F888140183C104FF4DCC75DFFF45E4395DE47CA68B45180145E0015DDCFF4DC80F8594FDFFFF8B451099F7
		FB8955F08945E885C00F8E450100008B45EC0FAFC38365DC008945D48B45E88945CC33C08945F88945FC89451C8945148945103945EC7E6085DB7E518B4D
		D88B45080FAFCB034D108D50020FAF4D18034DDC8BF08BF88945F403CA2BF22BFA2955F4895DC80FB6440E030145140FB60101451C0FB6440F010145FC8B
		45F40FB604080145F883C104FF4DC875D8FF45108B45103B45EC7CA08B4DD485C9740B8B451499F7F9894514EB048365140033F63BCE740B8B451C99F7F9
		89451CEB0389751C3BCE740B8B45FC99F7F98945FCEB038975FC3BCE740B8B45F899F7F98945F8EB038975F88975103975EC7E5585DB7E468B4DD88B450C
		0FAFCB034D108D50020FAF4D18034DDC8BF08BF803CA2BF22BFA2BC2895DC88A551488540E038A551C88118A55FC88540F018A55F888140183C104FF4DC8
		75DFFF45108B45103B45EC7CAB8BC3C1E0020145DCFF4DCC0F85CEFEFFFF8B4DEC33C08945F88945FC89451C8945148945103BC87E6C3945F07E5C8B4DD8
		8B75E80FAFCB034D100FAFF30FAF4D188B45088D500203CA8D0CB18BF08BF88945F48B45F02BF22BFA2955F48945C80FB6440E030145140FB60101451C0F
		B6440F010145FC8B45F40FB604010145F883C104FF4DC875D833C0FF45108B4DEC394D107C940FAF4DF03BC874068B451499F7F933F68945143BCE740B8B
		451C99F7F989451CEB0389751C3BCE740B8B45FC99F7F98945FCEB038975FC3BCE740B8B45F899F7F98945F8EB038975F88975083975EC7E63EB0233F639
		75F07E4F8B4DD88B75E80FAFCB034D080FAFF30FAF4D188B450C8D500203CA8D0CB18BF08BF82BF22BFA2BC28B55F08955108A551488540E038A551C8811
		8A55FC88540F018A55F888140883C104FF4D1075DFFF45088B45083B45EC7C9F5F5E33C05BC9C21800
		)
		else ; x64 machine code
		MCode_PixelateBitmap =
		(LTrim Join
		4489442418488954241048894C24085355565741544155415641574883EC28418BC1448B8C24980000004C8BDA99488BD941F7F9448BD0448BFA8954240C
		448994248800000085C00F8E9D020000418BC04533E4458BF299448924244C8954241041F7F933C9898C24980000008BEA89542404448BE889442408EB05
		4C8B5C24784585ED0F8E1A010000458BF1418BFD48897C2418450FAFF14533D233F633ED4533E44533ED4585C97E5B4C63BC2490000000418D040A410FAF
		C148984C8D441802498BD9498BD04D8BD90FB642010FB64AFF4403E80FB60203E90FB64AFE4883C2044403E003F149FFCB75DE4D03C748FFCB75D0488B7C
		24188B8C24980000004C8B5C2478418BC59941F7FE448BE8418BC49941F7FE448BE08BC59941F7FE8BE88BC69941F7FE8BF04585C97E4048639C24900000
		004103CA4D8BC1410FAFC94863C94A8D541902488BCA498BC144886901448821408869FF408871FE4883C10448FFC875E84803D349FFC875DA8B8C249800
		0000488B5C24704C8B5C24784183C20448FFCF48897C24180F850AFFFFFF8B6C2404448B2424448B6C24084C8B74241085ED0F840A01000033FF33DB4533
		DB4533D24533C04585C97E53488B74247085ED7E42438D0C04418BC50FAF8C2490000000410FAFC18D04814863C8488D5431028BCD0FB642014403D00FB6
		024883C2044403D80FB642FB03D80FB642FA03F848FFC975DE41FFC0453BC17CB28BCD410FAFC985C9740A418BC299F7F98BF0EB0233F685C9740B418BC3
		99F7F9448BD8EB034533DB85C9740A8BC399F7F9448BD0EB034533D285C9740A8BC799F7F9448BC0EB034533C033D24585C97E4D4C8B74247885ED7E3841
		8D0C14418BC50FAF8C2490000000410FAFC18D04814863C84A8D4431028BCD40887001448818448850FF448840FE4883C00448FFC975E8FFC2413BD17CBD
		4C8B7424108B8C2498000000038C2490000000488B5C24704503E149FFCE44892424898C24980000004C897424100F859EFDFFFF448B7C240C448B842480
		000000418BC09941F7F98BE8448BEA89942498000000896C240C85C00F8E3B010000448BAC2488000000418BCF448BF5410FAFC9898C248000000033FF33
		ED33F64533DB4533D24533C04585FF7E524585C97E40418BC5410FAFC14103C00FAF84249000000003C74898488D541802498BD90FB642014403D00FB602
		4883C2044403D80FB642FB03F00FB642FA03E848FFCB75DE488B5C247041FFC0453BC77CAE85C9740B418BC299F7F9448BE0EB034533E485C9740A418BC3
		99F7F98BD8EB0233DB85C9740A8BC699F7F9448BD8EB034533DB85C9740A8BC599F7F9448BD0EB034533D24533C04585FF7E4E488B4C24784585C97E3541
		8BC5410FAFC14103C00FAF84249000000003C74898488D540802498BC144886201881A44885AFF448852FE4883C20448FFC875E941FFC0453BC77CBE8B8C
		2480000000488B5C2470418BC1C1E00203F849FFCE0F85ECFEFFFF448BAC24980000008B6C240C448BA4248800000033FF33DB4533DB4533D24533C04585
		FF7E5A488B7424704585ED7E48418BCC8BC5410FAFC94103C80FAF8C2490000000410FAFC18D04814863C8488D543102418BCD0FB642014403D00FB60248
		83C2044403D80FB642FB03D80FB642FA03F848FFC975DE41FFC0453BC77CAB418BCF410FAFCD85C9740A418BC299F7F98BF0EB0233F685C9740B418BC399
		F7F9448BD8EB034533DB85C9740A8BC399F7F9448BD0EB034533D285C9740A8BC799F7F9448BC0EB034533C033D24585FF7E4E4585ED7E42418BCC8BC541
		0FAFC903CA0FAF8C2490000000410FAFC18D04814863C8488B442478488D440102418BCD40887001448818448850FF448840FE4883C00448FFC975E8FFC2
		413BD77CB233C04883C428415F415E415D415C5F5E5D5BC3
		)
		
		VarSetCapacity(PixelateBitmap, StrLen(MCode_PixelateBitmap)//2)
		Loop % StrLen(MCode_PixelateBitmap)//2		;%
			NumPut("0x" SubStr(MCode_PixelateBitmap, (2*A_Index)-1, 2), PixelateBitmap, A_Index-1, "UChar")
		DllCall("VirtualProtect", Ptr, &PixelateBitmap, Ptr, VarSetCapacity(PixelateBitmap), "uint", 0x40, A_PtrSize ? "UPtr*" : "UInt*", 0)
	}

	Gdip_GetImageDimensions(pBitmap, Width, Height)
	
	if (Width != Gdip_GetImageWidth(pBitmapOut) || Height != Gdip_GetImageHeight(pBitmapOut))
		return -1
	if (BlockSize > Width || BlockSize > Height)
		return -2

	E1 := Gdip_LockBits(pBitmap, 0, 0, Width, Height, Stride1, Scan01, BitmapData1)
	E2 := Gdip_LockBits(pBitmapOut, 0, 0, Width, Height, Stride2, Scan02, BitmapData2)
	if (E1 || E2)
		return -3

	E := DllCall(&PixelateBitmap, Ptr, Scan01, Ptr, Scan02, "int", Width, "int", Height, "int", Stride1, "int", BlockSize)
	
	Gdip_UnlockBits(pBitmap, BitmapData1), Gdip_UnlockBits(pBitmapOut, BitmapData2)
	return 0
}

;#####################################################################################

Gdip_ToARGB(A, R, G, B)
{
	return (A << 24) | (R << 16) | (G << 8) | B
}

;#####################################################################################

Gdip_FromARGB(ARGB, ByRef A, ByRef R, ByRef G, ByRef B)
{
	A := (0xff000000 & ARGB) >> 24
	R := (0x00ff0000 & ARGB) >> 16
	G := (0x0000ff00 & ARGB) >> 8
	B := 0x000000ff & ARGB
}

;#####################################################################################

Gdip_AFromARGB(ARGB)
{
	return (0xff000000 & ARGB) >> 24
}

;#####################################################################################

Gdip_RFromARGB(ARGB)
{
	return (0x00ff0000 & ARGB) >> 16
}

;#####################################################################################

Gdip_GFromARGB(ARGB)
{
	return (0x0000ff00 & ARGB) >> 8
}

;#####################################################################################

Gdip_BFromARGB(ARGB)
{
	return 0x000000ff & ARGB
}

;#####################################################################################

StrGetB(Address, Length=-1, Encoding=0)
{
	; Flexible parameter handling:
	if Length is not integer
	Encoding := Length,  Length := -1

	; Check for obvious errors.
	if (Address+0 < 1024)
		return

	; Ensure 'Encoding' contains a numeric identifier.
	if Encoding = UTF-16
		Encoding = 1200
	else if Encoding = UTF-8
		Encoding = 65001
	else if SubStr(Encoding,1,2)="CP"
		Encoding := SubStr(Encoding,3)

	if !Encoding ; "" or 0
	{
		; No conversion necessary, but we might not want the whole string.
		if (Length == -1)
			Length := DllCall("lstrlen", "uint", Address)
		VarSetCapacity(String, Length)
		DllCall("lstrcpyn", "str", String, "uint", Address, "int", Length + 1)
	}
	else if Encoding = 1200 ; UTF-16
	{
		char_count := DllCall("WideCharToMultiByte", "uint", 0, "uint", 0x400, "uint", Address, "int", Length, "uint", 0, "uint", 0, "uint", 0, "uint", 0)
		VarSetCapacity(String, char_count)
		DllCall("WideCharToMultiByte", "uint", 0, "uint", 0x400, "uint", Address, "int", Length, "str", String, "int", char_count, "uint", 0, "uint", 0)
	}
	else if Encoding is integer
	{
		; Convert from target encoding to UTF-16 then to the active code page.
		char_count := DllCall("MultiByteToWideChar", "uint", Encoding, "uint", 0, "uint", Address, "int", Length, "uint", 0, "int", 0)
		VarSetCapacity(String, char_count * 2)
		char_count := DllCall("MultiByteToWideChar", "uint", Encoding, "uint", 0, "uint", Address, "int", Length, "uint", &String, "int", char_count * 2)
		String := StrGetB(&String, char_count, 1200)
	}
	
	return String
}	 
