#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir   %A_ScriptDir% ; Ensures a consistent starting directory.

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

/*
****************************************************************************************************************************************************
************ Variable Setup *******************************************************************************
*****************************************************************************************
*/
Global Prefix_Number_Location_Check, First_Effectivity_Numbers, Title, sleepstill, Current_Monitor

Version_Number = 1.4 Beta
Effectivity_Macro :=  "Effectivity Macro V" Version_Number
Checkp=0
Sleepstill = 0
LoopCount = -1
ButtonLoop = 0
Searchcount = 0
Guicount = 1
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

inifile = c:\Serialmacro\config.ini
File_Install_Root_Folder = %A_ScriptDir%\Install_Files
File_Install_Work_Folder = C:\SerialMacro
Image_Red_Exclamation_Point = %File_Install_Work_Folder%\red_image.png
IMage_Actve_Add_Button = %File_Install_Work_Folder%\Active_plus.png
Image_Active_Apply_Button := %File_Install_Work_Folder%\orange_button.png

Unit_Test = 0 ; Set this to 1 to perform unit tests and logging. 

IfExist(A_ScriptDir "\Dev.txt")
	IfNotExist, %A_ScriptDir%\Install_Files
	{
	MsgBox, No Install file folder
	ExitApp
	}		

Result = Folder_Exist_Check("SerialMacro")
If Result = 0
	Folder_Create("SerialMacro")

Result := Folder_Exist_Check("SerialMacro\icons")
If Result = 0
	Folder_Create("SerialMacro\icons")

Result := File_Exist_Check("config.ini")
If Result = 0
	{
	File_Create("Config.ini")
	sleep(5)
	IniWrite, 20,  %inifile%,refreshrate,refreshrate
   Sleep()
	}

Load_ini_file(inifile)

If Sleep_Delay = Error
	{
	IniWrite, 3,  %inifile%,Sleep_Delay,Sleep_Delay
	}

If Refreshrate = Error
	{
	   IniWrite, 20,  %inifile%,refreshrate,refreshrate   
	}
  
Load_ini_file(inifile)
  
Install_Requied_Files(File_Install_Root_Folder,File_Install_Work_Folder)

Create_Tray_Menu()
Create_Main_GUI_Menu()

editfield := Temp_File_Read_Delete(File_Install_Root_Folder,"TempAdd.txt")
editfield2 := Temp_File_Read_Delete(File_Install_Root_Folder,"TempAdded.txt")
TotalPrefixes := Temp_File_Read_Delete(File_Install_Root_Folder,"TempAmount.txt")

If (editfield = null) || (editfield2=null) || (TotalPrefixes = null)
	{
	editfield = 
	editfield2 = 
	TotalPrefixes = 
	}

Msg_Box_Result := Update_Check(updatestatus)
If Msg_Box_Result = Yes
	Versioncheck()

Serials_GUI_Screen()
return

SerialsGUIscreen()
{
Global 

activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
 
   If totalprefixes < 1
   {
      TotalPrefixes = 0
   }
  
   gui 1:add, Edit, x10 y50 w390 h240  vEditField,%editfield%   
   gui 1:add, Edit, xp yp w390 h240 vEditField2,%editfield2%
   
   Gui 1:Add, Picture, x315 y310 w50 h50 +0x4000000  BackGroundTrans vStarting gstartmacro , C:\SerialMacro\Start.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vRunning, C:\SerialMacro\Running.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vpaused  gpausesub, C:\SerialMacro\Paused.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vStopped grestartmacro, C:\SerialMacro\Stopped.png
   Gui, 1:Add, Picture, x0 y0 w410 h400 +0x4000000 , C:\SerialMacro\background.png
   
   Gui 1:Add, Edit, xp+165 yp+343 w110 h20  vnextserialtoadd, %nextserialtoaddv%
   
   Gui 1:Add, Text, x5 y5 w300 h25 BackgroundTrans +Center vreloadprefixtext, There are a total of %TotalPrefixes% Effectivity to add to ACM
   
   Gui 1:add, Radio, xp+25 yp+25 w130 h20 BackGroundTrans vradio1 gradio1h, Effectivity to be added 
   Gui 1:add, Radio, xp+155 yp w140 h20 BackGroundTrans vradio2 gradio2h, Effectivity already added
   
   Gui 1:Add, Text, xp-170 Yp+265 W250 h13 BackGroundTrans vserialsentered, Number of Effectivity successfully added to ACM = %Serialcount%
   
   Gui 1:Add, Text, Xp Yp+15 w250 h13  BackGroundTrans , If macro is operating incorrectly, press Esc to reload
   Gui 1:Add, Text, xp yp+15 w250 h13  BackGroundTrans , Or press Pause Button on keyboard to Pause macro, Press Pause again to resume macro.
   Gui 1:Add, Text, xp yp+20 w145 h20  BackgroundTrans , Next Effectivity to add to ACM:   
   Gui 1:Menu, MyMenuBar
   Gui 1:Show,  x%amonx% y%amony% , %Effectivity_Macro%
   gui 1: +alwaysontop
   Editfield_Control("Editfield")

   ; IfExist  C:\SerialMacro\Tempcount.txt
   ; {
      ; FileRead, Serialcount,C:\SerialMacro\Tempcount.txt
      ; GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%	
      ; FileDelete, C:\SerialMacro\Tempcount.txt
   ; }
   Gui 1:Submit, NoHide
   Return
}

 Editfield_Control(Textbox)
 {
 If Textbox = Editfield
 {
   Guicontrol,hide, Editfield2,
   Guicontrol,show, Editfield,
   Guicontrol, Focus, Editfield
 }
 Else
 {
 Guicontrol,hide, Editfield,
  Guicontrol,show, Editfield2,   
   Guicontrol, Focus, Editfield2
  }
  return
 }
 
 
startmacro:
{
   skipbox = 0
   gosub, enterallserials
   return
}

   pausesub:
   {
      if A_IsPaused = 0
      {
         gosub, radio1h
         Gui 1: -AlwaysOnTop
		 Gui_Image_Show("Paused")
		 make the above function for the images
         Guicontrol,hide, Start
         Guicontrol,show, paused
         Guicontrol,hide, Stopped
         Guicontrol,hide, Running
         Gui, Submit, NoHide
         activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
         Titletext := "Press pause to unpause"
         Settimer, winmovemsgbox, 20
         Msgbox,262144,%Effectivity_Macro%, Macro is paused. Press pause to unpause,.1
         Settimer, winmovemsgbox, 20
         Msgbox,262144,%Effectivity_Macro%, Macro is paused. Press pause to unpause,.1
         Settimer, winmovemsgbox, 20
         Msgbox,262144,%Effectivity_Macro%, Macro is paused. Press pause to unpause,.1
         Settimer, winmovemsgbox, 20
         Msgbox,262144,%Effectivity_Macro%, Macro is paused. Press pause to unpause,.1
         Settimer, winmovemsgbox, 20
         Msgbox,262144,%Effectivity_Macro%, Macro is paused. Press pause to unpause,10
         Pause, toggle, 1 
         Return
      }else  {
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
   
   


Folder_Exist_Check(Folder)
	{
		Result := FileExist("C:\" Folder)
	return Result
	}

Folder_Create(Folder)
	{
		
	   FileCreateDir, C:\%Folder%
	   sleep()	   
	}
	return
	}

File_Exist_Check(File)
	{
		Result := FileExist("C:\SerialMacro\" File)
	return Result
	}

	File_Create(File)
	{
		FileAppend, "C:\SerialMacro\" %File%		
	return 
	}


Install_Requied_Files(File_Install_Root_Folder, File_Install_Work_Folder)
	{
	FileInstall, %File_Install_Root_Folder%\How to use Effectivity Macro.pdf, %File_Install_Work_Folder%\How to use Effectivity Macro.pdf,1
	FileInstall, %File_Install_Root_Folder%\icons\serial.ico, %File_Install_Work_Folder%\icons\serial.ico,1
	FileInstall, %File_Install_Root_Folder%\icons\paused.ico, %File_Install_Work_Folder%\icons\paused.ico,1
	FileInstall, %File_Install_Root_Folder%\serial macro\redimage.png, %File_Install_Work_Folder%\red_image.png,1	
	FileInstall, %File_Install_Root_Folder%\plus_sign.png, %File_Install_Work_Folder%\plus_sign.png,1
	FileInstall, %File_Install_Root_Folder%\active_plus.png, %File_Install_Work_Folder%\active_plus.png,1
	FileInstall, %File_Install_Root_Folder%\orange_button.png, %File_Install_Work_Folder%\orange_button.png,1
	FileInstall, %File_Install_Root_Folder%\paused.png, %File_Install_Work_Folder%\paused.png,1
	 FileInstall, %File_Install_Root_Folder%\start.png, %File_Install_Work_Folder%\start.png,1
	 FileInstall, %File_Install_Root_Folder%\Running.png, %File_Install_Work_Folder%\Running.png,1
	 FileInstall, %File_Install_Root_Folder%\Stopped.png, %File_Install_Work_Folder%\Stopped.png,1
	 FileInstall, %File_Install_Root_Folder%\background.png, %File_Install_Work_Folder%\background.png,1
	return
	}

Create_Tray_Menu()
{
Menu, Tray, NoStandard 
Menu, Tray, Add, How to use, HowTo
Menu Tray, Add, Check For update, Versioncheck
Menu, Tray, Add, Quit, Quitapp
return
}

Create_Main_GUI_Menu()
{
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
   Return
   }
   
Temp_File_Read_Delete(File_Install_Root_Folder,File_Name)
{
	IfExist(%File_Install_Root_Folder% "\" %File_Name% )
	FileRead, Variable_Store, %File_Install_Root_Folder%\%File_Name% 
	else
	Variable_Store = Null
	
	return Variable_Store
}

Move_Message_Box(Msg_box_type,Msg_box_title, Msg_box_text, Msg_box_Time := 2147483 )
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

   winmovemsgbox:
   {
      SetTimer, WinMoveMsgBox, OFF 
      WinMove, %Msg_box_text% , Amonx, Amony 
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
   
   
sleep(Amount := 1)
{
amount := amount * 100
Sleep %Amount%
Return
}

Updatechecker()
{
	global
    Today := A_Now		; Set to the current date first
      EnvSub, Today, %updatestatus% , Days 	; this does a date calc, in days
      If Today > %Updaterate%	; More than speficied days
      {
			Result := Move_Message_Box("4","Effectivity Macro Updater", "Would you like to check for a new update?" )
			IniWrite, %updatestatus%,  %inifile%,update,updaterate	
            IniWrite, %A_now%,  %inifile%, update,lastupdate
      return Result 
  }
  
Versioncheck()
{
 Progress,  w200, Updating..., Gathering Information, Effectivity Macro Updater
Progress, 0
   sleep(2)
   Versioncount = 0
   settimer, versiontimeout, 500
hwnd := create_checkgui()

   Progress,  w200, Updating..., Fetching Server Information, Effectivity Macro Updater
   Progress, 15
   
   DllCall("SetParent", "uint",  hwnd, "uint", ParentGUI)
   wb.Visible := True
   WinSet, Style, -0xC00000, ahk_id %hwnd%
   
   Progress, 25
   sleep(2)
   
   wb.navigate(Update_Check_URL) ; Update_Check_url is from Config FIle
   
   Progress,  w200, Updating...,Gathering Current Version From Server, Effectivity Macro Updater
   Progress, 50
   
   
   sleep(2)
   
   while wb.busy
   {
      sleep()
   }
   
   Progress,  w200,Updating..., Comparing Version Information, Effectivity Macro Updater
   Progress, 60
   Progress, Off


   ;msgbox, %Checkversion%
  
   Result = Check_Doc_Title()
   If Result = Not_Found
		{
		Progress,  w200,Updating..., Error Occured. Update Not Able To Complete, Effectivity Macro Updater
		Progress, 0
		 sleep(30)
		 Progress, off		
		 Move_Message_Box("0","Effectivity Macro Updater", "Error `n Cannot find Server" )
		 return
		  }
   
   
   If Checkversion >= %Version_Number%
   {
      Progress,  w200,Updating..., Macro is Up to date., Effectivity Macro Updater
      Progress, 100
      sleep(30)
      Progress, off
      settimer, versiontimeout, Off
      ;Msgbox,,Serial Macro Updater,Macro is Up to date.
   }
   
   If Checkversion > %Version_Number%
   {
	  settimer, versiontimeout, Off
      
      Result := Move_Message_Box("262148","Effectivity Macro Updater", " New update found. Would you like to open the Cat Box site to download the latest version?" )
      If Result =  yes 
            Run, %Program_Location_Link%
	}
      
	IniWrite, %updaterate%, %inifile%,update,updaterate	
	IniWrite, %A_now%,  %inifile%, update,lastupdate
   Gui,2:Destroy
   return
}

Check_Doc_Title()
{
   Loop, 3
{
	Winactivate, Serial version
   wingettitle, Google_Doc_Title, A
   ;Msgbox, Title is  %title% 
   StringGetPos, pos, %Google_Doc_Title%, #, 1	
   Update_Check_Version_Number := SubStr(Google_Doc_Title, pos+2)
   ;	msgbox, %string1%
   ;msgbox, check version is `n`n`n`n%checkversion%
   Checkversion := SubStr(Update_Check_Version_Number,1,3)
   
   If Google_Doc_Title != 
  {
  Progress, Off
  Result = Checkversion
	 break
 }
 else
	Result = Not_Found

}
 return Result
}

create_checkgui()
{
	global wb, ParentGUI
	
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
   return hwnd
}        

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
      Gui,2:Destroy
   }
   Return
}

  Load_ini_file(inifile)
               {
                  global
                  Ini_var_store_array:= Object()
                  Tab_placeholder  = 
                  loop,read,%inifile%
                  {
                     If A_LoopReadLine = 
                     continue
                     
                     if regexmatch(A_Loopreadline,"\[(.*)?]")
                     {
                        Section :=regexreplace(A_loopreadline,"(\[)(.*)?(])","$2")
                        StringReplace, Section,Section, %a_space%,,All
                        
                        If Tab_PLaceholder = 
                        {
                           Tab_placeholder := Section
                        }
                        Else
                           Tab_placeholder := Tab_placeholder "|" Section
                        
                        continue
                     }
                     
                     else if A_LoopReadLine != 
                     {
                        StringGetPos, keytemppos, A_LoopReadLine, =,
                        StringLeft, keytemp, A_LoopReadLine,%keytemppos%
                        StringReplace, keytemp,keytemp,%A_SPace%,,All
                        INIstoretemp := Keytemp ":" Section
                        Ini_var_store_array.Insert(INIstoretemp)
                        IniRead,%keytemp%, %inifile%, %Section%, %keytemp%
                     }}
                  
                  return
               }
               
               Write_ini_file(inifile)
               {
                  global
                  
                  for index, element in Ini_var_store_array 
                  {
                     StringSplit, INI_Write,element, `:
                     
                     Varname := INI_Write1
                     IniWrite ,% %INI_Write1%, %inifile%, %INI_Write2%, %INI_Write1%
                  } 
                  
                  return
               }
