#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#ErrorStdOut
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir   %A_ScriptDir% ; Ensures a consistent starting directory.
/*!
	Library: Effectivity Macro
		This macro Gets the user selected text from the CPI tool and enters it into ACM effectivity screen

	Author: Jarett Karnia
	Version: 2.0
*/




#KeyHistory 0
SetBatchLines, -1
SetDefaultMouseSpeed, 0
SetWinDelay, 0
SetControlDelay, 0
CoordMode, Mouse, Screen
;~ CoordMode, Pixel, Screen
#SingleInstance Force
++SetTitleMatchMode, 2
DetectHiddenWindows On
DetectHiddenText on
#InstallKeybdHook
#InstallMouseHook


Global Prefix_Number_Location_Check, First_Effectivity_Numbers, Title, Current_Monitor, Log_Events, Unit_test, File_Install_Work_Folder, Oneupserial, combineser, Active_ID, Image_Red_Exclamation_Point, At_home,Issues_Image, Ini_var_store_array

; below is for testing between home and work computer
If A_UserName = karnijs
At_home = 0
else
At_home = 1


/*

TODO ************************
*** complete*** Setup the Export to Excel. - Make it into a CSV file so that is works faster
Create more unit tests
Create Testing scripts



*/

;~ #include Unit_testing\Unit_testing.ahk  ; Uncomment this to run unit test modules, to narrow down what function is broken
/*
****************************************************************************************************************************************************
************ Variable Setup *******************************************************************************
*****************************************************************************************
*/


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
Radiobutton = 1

Unit_Test = 0 ; Set this to 1 to perform unit tests  and offline testing
Log_Events = 0 ;Set this to 1 to perform logging

inifile = config.ini
;~ File_Install_Root_Folder = C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files ; for testing at work
;~ File_Install_Root_Folder = E:\Git\Effectivity_Macro\1.4\Install_Files ; for when testing at home


File_Install_Work_Folder = C:\SerialMacro
Configuration_File_Location = %File_Install_Work_Folder%\%inifile%
File_install_Image_Folder = %File_Install_Work_Folder%\images
File_install_Icon_Folder = %File_Install_Work_Folder%\icons

Image_Red_Exclamation_Point = %File_install_Image_Folder%\red_image.png
IMage_Actve_Add_Button = %File_install_Image_Folder%\Active_plus.png
Image_Active_Apply_Button = %File_install_Image_Folder%\orange_button.png
Issues_Image = %File_install_Image_Folder%\Issues_Image.png



Result := Folder_Exist_Check(File_Install_Work_Folder)
If Result contains Folder_Not_Exist
Folder_Create(File_Install_Work_Folder)

Result := Folder_Exist_Check(File_install_Image_Folder)
If Result contains  Folder_Not_Exist
Folder_Create(File_install_Image_Folder)

Result := Folder_Exist_Check(File_install_Icon_Folder)
If Result contains Folder_Not_Exist
Folder_Create(File_install_Icon_Folder)

Result := Config_File_Check(Configuration_File_Location)

If Result  contains File_Not_Exist
{
	Config_File_Create(Configuration_File_Location, At_home)
	sleep(5)
	IniWrite, 20,  %Configuration_File_Location%,refreshrate,refreshrate
		IniWrite, 3,  %Configuration_File_Location%,Sleep_Delay,Sleep_Delay
	Sleep()
}

If Sleep_Delay = Error
{

	IniWrite, 3,  %Configuration_File_Location%,Sleep_Delay,Sleep_Delay
	Load_ini_file(Configuration_File_Location)
}

If Refreshrate = Error
{
	IniWrite, 20,  %Configuration_File_Location%,refreshrate,refreshrate
	Load_ini_file(Configuration_File_Location)
}

Load_ini_file(Configuration_File_Location)

Result := Install_Requied_Files_Root(File_Install_Work_Folder, At_home)
If (Result)
	Move_Message_Box("0","Error","Error with installing How To PDF File. `n`nThis will not effect Program Operation")

Result := Install_Requied_Files_Icons(File_Install_Work_Folder, At_home)
If (Result)
	Move_Message_Box("0","Error","Error with installing Program Icons. `n`nThis will not Effect Program Operation")

Install_Requied_Files_Images(File_Install_Work_Folder, At_home)
If (Result)
{
	Move_Message_Box("0","Error","Error with installing Images needed for screen searching.`n`n Program will not run without these files.`n`n Please restart Program and if error occurs agian, Contact Jarett Karnia for assisstance")
	If (!unit_test)
		ExitApp
}

Create_Tray_Menu()
Create_Main_GUI_Menu()

editfield := Temp_File_Read(File_Install_Work_Folder,"TempAdd.txt")

Temp_File_Delete(File_Install_Work_Folder,"TempAdd.txt")

editfield2 := Temp_File_Read(File_Install_Work_Folder,"TempAdded.txt")
Temp_File_Delete(File_Install_Work_Folder,"TempAdded.txt")

TotalPrefixes := Temp_File_Read(File_Install_Work_Folder,"TempAmount.txt")
Temp_File_Delete(File_Install_Work_Folder,"TempAmount.txt")

If (editfield = "null") || (editfield2= "null") || (TotalPrefixes = "null")
{
	editfield =
	editfield2 =
	TotalPrefixes =
}



;~ Versioncheck()
	


Serials_GUI_Screen(editfield, editfield2, TotalPrefixes)

If (Unit_test)
	Checkvalues(Prefix_Store, First_Number_Set,  Second_Number_Set, "1")
return

/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Hotkeys \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/
#`::ListLines
+#~::ListVars

   Insert::
   Pause::
   {
      if A_IsPaused = 0
      {
         Gui 1: -AlwaysOnTop
         Gui, Submit, NoHide

  Loop, 4
	Move_Message_Box("262144", Effectivity_Macro,"Macro is paused. Press pause to unpause", ".1")

	Move_Message_Box("262144", Effectivity_Macro,"Macro is paused. Press pause to unpause", "10")

    Pausescript()
         Return
      }else  {
       gosub, radio_button
         Gui 1: +AlwaysOnTop
UnPausescript()
         Gui, Submit, NoHide
      }
      return
   }

^q::
{
   Exit_Program()
   Return
}


ESC::
{
breakloop=1
   ListLines
   Pause,on
return
}

;Sets the hotkey for Ctrl + 1 or Ctrl + numpad 1
$^Numpad1::
$^1::
{
Copy_text_and_Format()
return
}

$^Numpad2::
$^2::
{
Start_Macro()
return
}

#If Winactive("ahk_class TTAFrameXClass") or WinActive(Effectivity_Macro)
~Esc::
{
   Gosub,Stop_Macro
   ListLines
   Pause,on
   Return
}

#if winactive(Effectivity_Macro)


F1::
{
  HowTo()
   Return
}

#if winactive ; stops the requirement for only the macro screen or acm



Stop_Macro:
{
Result := Move_Message_Box("262148", Effectivity_Macro, "The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure that you want to stop the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.")
   If result =  yes
   {
      Stopactcheck = 1
      Gui 1: -AlwaysOnTop
	  Gui_Image_Show("Stop")
      Gui, Submit, NoHide
      Send {Shift Up}{Ctrl Up}
      breakloop = 1
      Exit
   }
   Else
      Return
}
/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Below are the functions from the Autorun section Before it gets to Serials_GUI_Screen()i \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

	Folder_Exist_Check(Folder) ; unit  && Documentation
	{
		Result := FileExist(Folder)
		If Result =
		Result = Folder_Not_Exist
		else
			Result = Folder_Exist

		Debug_Log_Event("Folder_Exist_Check() .... " Folder "..... Result is " Result)

		return  Folder " - " Result

/*!
	Function: Folder_Exist_Check(Folder)
			Checks to ensure that a folder exists on the users computer

	Parameters:
		Folder - This should contain the full file path of the folder
					 >  Folder = C:\SerialMacro
					>  Folder_Exist_Check(Folder)

	Remarks:
		Uses the AutoHotkey built in function of `FileExist()` to determing is the file is on the computer

	Returns:
		The Name of the *Folder* Variable - the *result*  Variable
			> Folder := C:\SerialMacro
		> Result := Folder_Exist_Check(Folder)
		Result would be `C:\SerialMacro - Folder_Exist` if the folder **IS** on the C Drive
		Result would be `C:\SerialMacro - Folder_Not_Exist` if the folder was **NOT** on the C Drive

	Extra:
		### Additional Information
		For more information on the `FileExist()` function click on the link (Internet Connection Required) below:  
		[FileExist()](https://autohotkey.com/docs/commands/FileExist.htm)
*/
	}

	Folder_Create(Folder) ; unit && Documentation
	{
		FileCreateDir, %Folder%

		Debug_Log_Event("Folder_Create() .... " Folder "..... Result is " Result)
		sleep(5)
		return ErrorLevel

/*!
	Function:  Folder_Create(Folder)
			Creates a folder the users computer

	Parameters:
		Folder - This should contain the full file path of the folder
					 >  Folder = C:\SerialMacro
					>  Folder_Create(Folder)

	Remarks:
		Uses the AutoHotkey built in function of `FileCreateDir` to create a folder

	Returns:
		The Errorlevel for the `FileCreateDir`
		> Folder := C:\SerialMacro
		> Result := Folder_Create(Folder)
		*Result* is 0 if the folder was created
		*Result* is 1 if the folder was not able to be created

	Extra:
		### Additional Information
		For more information on the `FileCreateDir` function click on the link (Internet Connection Required) below:  
		[FileCreateDir](https://autohotkey.com/docs/commands/FileCreateDir.htm)


*/
	}


	Config_File_Check(File) ;unit && Documentation
	{
		Result := FileExist(File)
		If Result =
		Result = File_Not_Exist
		else
			Result = File_Exist
		Debug_Log_Event("File_Exist_Check() ......C:\SerialMacro\" File "..... Result is " Result)
		return File " - "  Result

/*!
	Function:  Config_File_Check(File)
			checks the  existance of  the config file

	Parameters:
		File - This should contain the full file path of the file
					 >  Config_File = C:\SerialMacro\Config.ini
					>  Config_File_Check(Config_File)

	Remarks:
		Uses the AutoHotkey built in function of `FileExist()` to check for the config file

	Returns:
		The Name of the *File* Variable - the *result*  Variable
		>  Config_File = C:\SerialMacro\Config.ini
		>  REsult :=- Config_File_Check(Config_File)
		Result would be `C:\SerialMacro\Config.ini - File_Exist` if the file **IS** at the location specified in the *File* variable
		Result would be `C:\SerialMacro\Config.ini - File_Not_Exist` if the folder was **NOT**  location specified in the *File* variable

	Extra:
		### Additional Information
			For more information on the `FileExist()` function click on the link (Internet Connection Required) below:  
			[FileExist()](https://autohotkey.com/docs/commands/FileExist.htm)
*/


	}

	Config_File_Create(File, At_home:= 0) ; unit && Documentation
	{
			If (at_home)
					FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\Config.ini, %File%,1
			else
					FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\Config.ini, %File%,1
		Debug_Log_Event("File_Create() ......" File)
		return ErrorLevel
/*!
	Function: 		Config_File_Create(File [, At_home:= 0])
			Creates a config file on  the users computer at the location specified in the *File* variable

	Parameters:
		File - This should contain the full file path of the Config file
					 >  File = C:\SerialMacro\config.ini
					>  Config_File_Create(File)  
		At_Home (Optional)-   0: **Default** if nothing is in that location on the function
				> Config_File_Create(File)
				1:  Function will use the alternate Fileinstall location
				> Config_File_Create(File, "1")


	Remarks:
		Uses the AutoHotkey built in function of `FileInstall` to combine a file with the compiled version of the program.
		The 'At_home' optional variable is here because I got tired of having errors when testing at home, and instead of having to change the orginal file location from
		> C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\Config.ini
		to
		> E:\Git\Effectivity_Macro\1.4\Install_Files\Config.ini
		when I was testing from home.
		I made it so that the program knows which computer I am on automatically. In the beginning of the program, I have the following code:
		> If A_UserName = karnijs
		> At_home = 0
		> else
		> At_home = 1
		If I am on the work computer it will detect my username, or else it will be the home computer.   I did this because you cannot have a variable for the initlal location of the file for `FileInstall`. This is because AutohotKey does not know where the file is to install on compile if there is a variable in that location. Just saved me some time in the long run.

	Returns:
		The Errorlevel for the `Fileinstall` function
			> Folder := C:\SerialMacro
		> Result := Folder_Create(Folder)
		*Result* is 0 if the folder was created
		*Result* is 1 if the folder was not able to be created

	Extra:
		### Additional Information
			For more information on the `FileInstall` function click on the link (Internet Connection Required) below:
			[FileInstall](https://autohotkey.com/docs/commands/FileInstall.htm)

*/
	}



	Install_Requied_Files_Root( File_Install_Work_Folder, At_home:= 0) ; no unit testing as functions have built in error checking && Documentation
	{
		If (at_home)
					FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\How to use Effectivity Macro.pdf, %File_Install_Work_Folder%\How to use Effectivity Macro.pdf,1
					else
					FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\How to use Effectivity Macro.pdf, %File_Install_Work_Folder%\How to use Effectivity Macro.pdf,1
		return errorlevel

/*!
	Function: Install_Requied_Files_Root( File_Install_Work_Folder [, At_home:= 0])
			Creates the How To Use PDF  file on  the users computer at the folder specifiec by the *File_Install_Work_Folder*

	Parameters:
		File_Install_Work_Folder - This should contain the full file path of the Root script working folder
					 >  File_Install_Work_Folder = C:\SerialMacro
					>  Function: Install_Requied_Files_Root( File_Install_Work_Folder)
		At_Home (Optional)-   0: **Default** tf nothing is in that location on the function
				> Install_Requied_Files_Root( File_Install_Work_Folder)
				1:  Function will use the alternate Fileinstall location
				> Install_Requied_Files_Root( File_Install_Work_Folder, "1")


	Remarks:
		Uses the AutoHotkey built in function of `FileInstall` to combine a file with the compiled version of the program.
		The 'At_home' optional variable is here because I got tired of having errors when testing at home, and instead of having to change the orginal file location from
		> C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files
		to
		> E:\Git\Effectivity_Macro\1.4\Install_Files\
		when I was testing from home.
		I made it so that the program knows which computer I am on automatically. In the beginning of the program, I have the following code:
		> If A_UserName = karnijs
		> At_home = 0
		> else
		> At_home = 1
		If I am on the work computer it will detect my username, or else it will be the home computer.   I did this because you cannot have a variable for the initlal location of the file for `FileInstall`. This is because AutohotKey does not know where the file is to install on compile if there is a variable in that location. Just saved me some time in the long run.


	Returns:
		The Errorlevel for the `Fileinstall` function
			> File_Install_Work_Folder = C:\SerialMacro
		> Result := Install_Requied_Files_Root( File_Install_Work_Folder)
		*Result* is 0 if the file was created
		*Result* is 1 if the file was not able to be created

	Extra:
		### Additional Information
			For more information on the `FileInstall` function click on the link (Internet Connection Required) below:
			[FileInstall](https://autohotkey.com/docs/commands/FileInstall.htm)

*/
	}

	Install_Requied_Files_Icons( File_Install_Work_Folder, at_home := 0)  ; no unit testing as functions have built in error checking && Documentation
	{
		Problems = 0
		If (at_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\icons\serial.ico, %File_Install_Work_Folder%\icons\serial.ico,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\icons\serial.ico, %File_Install_Work_Folder%\icons\serial.ico,1
		If (Errorlevel)
			Problems = 1

		if (At_home)
				FileInstall, E:\Git\Effectivity_Macro\1.4\Install_Files\icons\paused.ico, %File_Install_Work_Folder%\icons\paused.ico,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\icons\paused.ico, %File_Install_Work_Folder%\icons\paused.ico,1

		If (Errorlevel)
			Problems = 1

		return Problems

/*!
	Function: Install_Requied_Files_Icons( File_Install_Work_Folder [, at_home := 0])
			Creates the icon  files on  the users computer at the folder specifiec by the *File_Install_Work_Folder*

	Parameters:
		File_Install_Work_Folder - This should contain the full file path of the Root script working folder
					 >  File_Install_Work_Folder = C:\SerialMacro
					>  Function: Install_Requied_Files_Icons( File_Install_Work_Folder)
		At_Home (Optional)-   0: **Default** tf nothing is in that location on the function
				> Install_Requied_Files_Icons( File_Install_Work_Folder)
				1:  Function will use the alternate Fileinstall location
				> Install_Requied_Files_Icons( File_Install_Work_Folder, "1")


	Remarks:
		Uses the AutoHotkey built in function of `FileInstall` to combine a file with the compiled version of the program.
		The 'At_home' optional variable is here because I got tired of having errors when testing at home, and instead of having to change the orginal file location from
		>  C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\icons\serial.ico
		>  C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\icons\paused.ico
		to
		> E:\Git\Effectivity_Macro\1.4\Install_Files\icons\serial.ico
		> E:\Git\Effectivity_Macro\1.4\Install_Files\icons\paused.ico
		when I was testing from home.
		I made it so that the program knows which computer I am on automatically. In the beginning of the program, I have the following code:
		> If A_UserName = karnijs
		> At_home = 0
		> else
		> At_home = 1
		If I am on the work computer it will detect my username, or else it will be the home computer. I did this because you cannot have a variable for the initlal location of the file for `FileInstall`. This is because AutohotKey does not know where the file is to install on compile if there is a variable in that location. Just saved me some time in the long run.


	Returns:
		The *Problems* variable
		If there is an issue and one of the files cannot be installed, the Errorlevel of the `Fileinstall` will trigger a `Problems = 1`
			> File_Install_Work_Folder = C:\SerialMacro
		> Result := Install_Requied_Files_Icons( File_Install_Work_Folder)
		*Result* is 0 if the files were created
		*Result* is 1 if  at least one of the files was not able to be created

	Extra:
		### Additional Information
			For more information on the `FileInstall` function click on the link (Internet Connection Required) below:
			[FileInstall](https://autohotkey.com/docs/commands/FileInstall.htm)

*/
	}

	Install_Requied_Files_Images( File_Install_Work_Folder, at_home := 0)  ; no unit testing as functions have built in error checking && Documentation
	{
		Problems = 0
		if (At_home)
					FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\red_image.png, %File_Install_Work_Folder%\images\red_image.png,1
					else
					FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\red_image.png, %File_Install_Work_Folder%\images\red_image.png,1
		If (Errorlevel)
			Problems = 1

		If (At_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\active_plus.png, %File_Install_Work_Folder%\images\active_plus.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\active_plus.png, %File_Install_Work_Folder%\images\active_plus.png,1
		If (Errorlevel)
			Problems = 1
		If (At_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\orange_button.png, %File_Install_Work_Folder%\images\orange_button.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\orange_button.png, %File_Install_Work_Folder%\images\orange_button.png,1
		If (Errorlevel)
			Problems = 1
		If (At_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\paused.png, %File_Install_Work_Folder%\images\paused.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\paused.png, %File_Install_Work_Folder%\images\paused.png,1
		If (Errorlevel)
			Problems = 1
		If (At_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\start.png, %File_Install_Work_Folder%\images\start.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\start.png, %File_Install_Work_Folder%\images\start.png,1
		If (Errorlevel)
			Problems = 1
		If (At_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\Running.png, %File_Install_Work_Folder%\images\Running.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Running.png, %File_Install_Work_Folder%\images\Running.png,1
		If (Errorlevel)
			Problems = 1
		If (At_home)
		FileInstall, E:\Git\Effectivity_Macro\1.4\Install_Files\images\Stopped.png, %File_Install_Work_Folder%\images\Stopped.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Stopped.png, %File_Install_Work_Folder%\images\Stopped.png,1
		If (Errorlevel)
			Problems = 1
		If (At_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\background.png, %File_Install_Work_Folder%\images\background.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\background.png, %File_Install_Work_Folder%\images\background.png,1
		If (Errorlevel)
			Problems = 1

			If (At_home)
		FileInstall,E:\Git\Effectivity_Macro\1.4\Install_Files\images\Issues_Image.png, %File_Install_Work_Folder%\images\Issues_Image.png,1
		else
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Issues_Image.png, %File_Install_Work_Folder%\images\Issues_Image.png,1
		If (Errorlevel)
			Problems = 1

		return Problems

/*!
	Function: 	Install_Requied_Files_Images( File_Install_Work_Folder [, at_home := 0])
			Creates the image  files on  the users computer at the folder specifiec by the *File_Install_Work_Folder*

	Parameters:
		File_Install_Work_Folder - This should contain the full file path of the Root script working folder
					 >  File_Install_Work_Folder = C:\SerialMacro
					>  Install_Requied_Files_Images( File_Install_Work_Folder)
		At_Home (Optional)-   0: **Default** tf nothing is in that location on the function
				> Install_Requied_Files_Images( File_Install_Work_Folder)
				1:  Function will use the alternate Fileinstall location
				> Install_Requied_Files_Images( File_Install_Work_Folder, "1")


	Remarks:
		Uses the AutoHotkey built in function of `FileInstall` to combine a file with the compiled version of the program.
		The 'At_home' optional variable is here because I got tired of having errors when testing at home, and instead of having to change the orginal file location from
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\red_image.png
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\active_plus.png
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\orange_button.png
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\paused.png
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Running.png
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Stopped.png
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\background.png
		>   C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Issues_Image.png
		to
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\red_image.png
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\active_plus.png
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\orange_button.png
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\paused.png
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\Running.png
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\Stopped.png
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\background.png
		> E:\Git\Effectivity_Macro\1.4\Install_Files\images\Issues_Image.png
		when I was testing from home.
		I made it so that the program knows which computer I am on automatically. In the beginning of the program, I have the following code:
		> If A_UserName = karnijs
		> At_home = 0
		> else
		> At_home = 1
		If I am on the work computer it will detect my username, or else it will be the home computer. I did this because you cannot have a variable for the initlal location of the file for `FileInstall`. This is because AutohotKey does not know where the file is to install on compile if there is a variable in that location. Just saved me some time in the long run.


	Returns:
		The *Problems* variable
		If there is an issue and one of the files cannot be installed, the Errorlevel of the `Fileinstall` will trigger a `Problems = 1`
			> File_Install_Work_Folder = C:\SerialMacro
		> Result := Install_Requied_Files_Images( File_Install_Work_Folder)
		*Result* is 0 if the files were created
		*Result* is 1 if  at least one of the files was not able to be created

	Extra:
		### Additional Information
			For more information on the `FileInstall` function click on the link (Internet Connection Required) below:  
			[FileInstall](https://autohotkey.com/docs/commands/FileInstall.htm)

*/
	}


				Load_ini_file(Configuration_File_Location) ; unit && Documentation
				{
					global

					Ini_var_store_array:= Object()
					Tab_placeholder  =
					loop,read,%Configuration_File_Location%
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
							IniRead,%keytemp%, %Configuration_File_Location%, %Section%, %keytemp%
						}}

					return

/*!
	Function: 	Load_ini_file(Configuration_File_Location)
			Loads the configuration Ini file from the  *Configuration_File_Location*

	Parameters:
		Configuration_File_Location - This should contain the full file path of the Root script working folder
					 >  Configuration_File_Location = C:\SerialMacro\Config.ini
					>  Load_ini_file(Configuration_File_Location)

	Remarks:
		Uses the AutoHotkey built in function of `IniRead` to read the config file.
		The function stores the variable names into an array for later retrievial by the `Write_ini_file()` function

	Returns:
		There is no returned variable. The Funciton is Global, which makes all the variables it stores global variables, which makes the config file contents accessble to all functions

	Extra:
		### Additional Information
			For more information on the `IniRead` function click on the link (Internet Connection Required) below:  
			[IniRead](https://autohotkey.com/docs/commands/IniRead.htm)

*/
				}

				Write_ini_file(Configuration_File_Location) ; unit && Documentation
				{
					global

					for index, element in Ini_var_store_array
					{
					StringSplit, INI_Write,element, `:

					Varname := INI_Write1
					IniWrite ,% %INI_Write1%, %Configuration_File_Location%, %INI_Write2%, %INI_Write1%
				}
			return

/*!
	Function: 	Write_ini_file(Configuration_File_Location)
			Writes the configuration variables to the configuration Ini file at  the  *Configuration_File_Location*

	Parameters:
		Configuration_File_Location - This should contain the full file path of the Root script working folder
					 >  Configuration_File_Location = C:\SerialMacro\Config.ini
					>  Write_ini_file(Configuration_File_Location)

	Remarks:
		Uses the AutoHotkey built in function of `IniWrite` to read the config file.
		The  highlighted  (green) **%** in the below code is there so written value to the `config.ini`  file ithe actual value of the variable and not the name of the variable from  the `ini_store_Array'
		`IniWrite ,*%* %INI_Write1%, %Configuration_File_Location%, %INI_Write2%, %INI_Write1%`



	Returns:
		There is no returned variable. The Funciton is Global, which makes all the variables it stores global variables, which makes the config file contents accessble to all functions

	Extra:
		### Additional Information
			For more information on the `IniRead` function click on the link (Internet Connection Required) below:  
			[IniWrite](https://autohotkey.com/docs/commands/IniWrite.htm)

*/
		}

		Debug_Log_Event(Event) ; no unit tesing && Documentation
		{
			global Log_Events

			If (Log_Events)
			{
			OutputDebug, %Event%
			Sleep(.5)

	}
	return


/*!
	Function: Debug_Log_Event(Event)
			Whe using the Debugger in the SciTE4AutoHotkey , this function will display the *Event* that is passed to it in a window. If `Log_events` = 1

	Parameters:
		Event - This is the text that will diaplay in the debugger window.

	Remarks:
		Uses the AutoHotkey built in function of `OutputDebug`  
		**NOTE:** the `Sleep(.5)` needs to be there or else the debugger window may not keep up with the even displays and crash.

	Returns:
		There is no returned variable.

	Extra:
		### Additional Information
			For more information on the `OutputDebug` function click on the link (Internet Connection Required) below:  
			[OutputDebug](https://autohotkey.com/docs/commands/OutputDebug.htm)

*/
}


/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Below is the Create_Tray_Menu funciton along with some of the functions it calls to .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

	Create_Tray_Menu() ; no unit tests && Documentation
	{
		Menu, Tray, NoStandard
		Menu, Tray, Add, How to use, HowTo
		Menu Tray, Add, Check For update, Versioncheck
		Menu, Tray, Add, Quit, Quitapp
		return

/*!
	Function: Create_Tray_Menu()
			Creates the system tray menu on the taskbar  
		
	Returns:
		There is no returned variable.

	Extra:
		### Additional Information
			For more information on the `Menu` function click on the link (Internet Connection Required) below:    
			[Menu](https://autohotkey.com/docs/commands/Menu.htm)

*/
	}


	Howto() ; no unit testing && Documentation
	{
		splashtexton,,Effectivity Macro, Loading PDF
		Run, C:\SerialMacro\How to use Effectivity Macro.pdf
		sleep(20)
		SplashTextOff
		return

/*!
	Function: Howto()  
		Opens the how to use PDF  

	Returns:
		There is no returned variable.

	Extra:
		### Additional Information
			For more information on the `Splashtext` function  or the `run` function click on the link (Internet Connection Required) below:  
			[Splashtext](https://autohotkey.com/docs/commands/Splashtext.htm)  
			[Run](https://autohotkey.com/docs/commands/Run.htm)

*/
	}

	Quitapp: ; no unit testing
	{
		Result := 	Move_Message_Box("262148","Quit " Effectivity_Macro, "Are you sure you want to quit?")

		If result =  Yes
		{
			Stopactcheck = 1
			Gui 1: -AlwaysOnTop
			Gui_Image_Show("Stop") ; Options are Start, Paused, Running, Stopped
			Gui, Submit, NoHide
			Send {Shift Up}{Ctrl Up}
			breakloop = 1
			ExitApp
		}
		Return
	}

/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. BEgin_Macro() from the Ctrl 1 grab text hotkey .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

Copy_text_and_Format() ; no unit test needed as it all the other functions are tested individually.
{
	Formatted_Text := Format_Serial_Functions(,Unit_Test) ; Goes to the Formatserials subroutine
	Sort, Formatted_Text ; Sorts the prefixes in order
	Guicontrol,,Radio1,1
	Gui,Submit,NoHide
	gosub, radio_button
	Debug_Log_Event("Formatted text is " Formatted_Text)

	Formatted_Serial_Array := Object()

	Formatted_Serial_Array := Put_Formatted_Serials_into_Array(Formatted_Text)

	/* for testing********
	*/

	Checked := SerialbreakquestionGUI() ; Goes to the Serialsgui.ahk and into the SerialbreakquestionGUI subroutine
	;~ combine = 0
	;~ Oneupserial = 0

	/*
	Stop for testing
	*/

	If (combine = "1") || (Oneupserial = "1")
	{
		Combined_Serial_Array := Combineserials(Formatted_Serial_Array) ;goes to the combine Serials subroutine

		Prefix_Count :=  Combined_Serial_Array.Length()

		If Oneupserial = 1
		Combined_Serial_Array := One_Up_All(Combined_Serial_Array)

		Editfield := Extract_Serial_Array(Combined_Serial_Array)
StringReplace, Editfield, Editfield, `,,,All
		Guicontrol,1:, Editfield, %Editfield% - - - - - - - - - - - - - - - - - - - - - - - - - - ; Sets the listbox on teh GUi screen to the editfieldcombine vaariable and adds a newline
	}else  {
		Prefix_Count :=  Formatted_Text_Serial_Count(Formatted_Text)
	StringReplace, Formatted_Text, Formatted_Text, `,,,All
		Guicontrol,1:, Editfield, %Formatted_Text% - - - - - - - - - - - - - - - - - - - - - - - - - - ; Sets the listbox on teh GUi screen to the editfieldcombine vaariable and adds a newline
	}
	totalprefixes = %Prefix_Count% ; Sets the totalprefixes variables to the Prefixcombinecount variable
		Guicontrol,, reloadprefixtext,%totalprefixes% ; Changes the valuse in the main GUI screen
	Winactivate, Effectivity Macro ; Make the Main GUi window  Active
	Guicontrol, Focus, Editfield ; Puts the cursor in the Editfield in teh Gui window
	send {Ctrl Down}{Home}{Ctrl Up} ; sends keystrokes to move the cursor to the top of the listbox
Gui, Submit, NoHide ; Updates the Gui screen
Gui_Image_Show("Start")
if (checked)
gosub, Export_Serials
	return
}

/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
Formatting functions
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/
Format_Serial_Functions(Fullstring := "", Unit_test := 0) ; unit
{
	Clear_Format_Variables(unit_test)
	newline = `n
	sleep()
	If (Unit_test) && (FullString = "") ; For testing
	Fullstring := "621s (SN: TRD00123, TRD00124-00165, TRD00001-00002, CHU00001-00002, CHU00003-00005)"
	Else if(!Unit_Test)
		FullString := Copy_selected_Text()


	If FullString = No_Text_Selected
	{
		Move_Message_Box("0","Error","Please ensure that text is selected before pressing Ctrl + 1.")
		Exit
	}
	Format_Removed_Text := Remove_Formatting(Fullstring) ; Goes to the Remove_Formatting funciton and stores the completed result into the Format_Removed_Text varialbe
		Debug_Log_Event("Format_Serials() Format_Removed_Text is " Format_Removed_Text)
;~ MsgBox, % "Format removed is " Format_Removed_Text
	PreFormatted_Text := PreFormat_Text(Format_Removed_Text) ; Goes to the PreFormat_Text function and stors the completed result into the Preformatted_text variable
	Debug_Log_Event("Format_Serials() PreFormatted_Text is " PreFormatted_Text)
	PreFormatted_Text := Check_For_Single_Serials(PreFormatted_Text)
Debug_Log_Event("Check_For_Single_Serials() PreFormatted_Text is " PreFormatted_Text)

	PreFormatted_Text := Prefix_Alone_Check_And_Add_One_UP(PreFormatted_Text) ;  Take info from the PreFormatted_Text variable and checks to see of it it just the serial with no numbers attached to it. If is, then adds 00001-99999 to prefix. Also changes the Parseclip variable to the Number of non combined Serials
	Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP() PreFormatted_Text is " PreFormatted_Text)
PreFormatted_Text := add_digits(PreFormatted_Text)
Debug_Log_Event("add_digits() PreFormatted_Text is " PreFormatted_Text)
	return PreFormatted_Text
}



Put_Formatted_Serials_into_Array(Formatted_Text) ; unit
{
	Formatted_Array := Object()
	Loop, Parse, Formatted_Text, `r`n
	{
		If A_LoopField =
		continue
		Formatted_Array.Insert(A_LoopField)
	}
	return Formatted_Array
}

Combineserials(Formatted_Serial_Array) ; unit
{
	Local Prefix_Combine_array := Object()


	For index, element in Formatted_Serial_Array
	{
		;~ MsgBox, Formatted serial is %element%
		Prefix_Extract := Extract_Prefix(element)
		If (Prefix_Extract = "`," or Prefix_Extract ="" or Prefix_Extract = "`n" or Prefix_Extract = "`r") ; checks to see if the Prefix_Store variable is a comma
		{
			Debug_Log_Event("Combine_Serials() Prefix_Extract is " Prefix_Extract)
			Debug_Log_Event("Combine_Serials() Continue" )
			Prefix_Extract = ; sets the Prefix_Store variable to nothing
			Second_Number_set =  ; sets the Second_Number_Set variable to nothirng
			Continue ; skips over the rest of the loop and starts at the top of the parse loop
		}

		First_Number_Set := Extract_First_Set_Of_Serial_Number(element)
		Middle_Char := Extract_Serial_Dividing_Char(element)

		If Middle_Char = `- ; checks if Middle_Char variable is a hyphen
		Second_Number_set := Extract_Second_Set_Of_Serial_Number(element)

		else If Middle_Char = `, ; if the Middle_Char variable is a comma
		{
			Middle_Char = `- ; makes the Middle_Char variable a hyphen
			Second_Number_set = %First_Number_Set%
		}
		Prefix_Combine_array := (Checkvalues(Prefix_Extract,First_Number_Set,Second_Number_set)) ; goes to the Checkvalues subroutine

		;~ For index, element in Prefix_Combine_array
		;~ MsgBox, % element " Is element`n index is " index

		Debug_Log_Event("Combine_Serials() Prefix_Extract is "  Prefix_Extract)
		Debug_Log_Event("Combine_Serials() First_Number_Set is "  First_Number_Set)
		Debug_Log_Event("Combine_Serials() Middle_Char is "  Middle_Char)
		Debug_Log_Event("Combine_Serials() Second_Number_set is "  Second_Number_set)
	}
	Return Prefix_Combine_array
}










Extract_Serial_Array(Combined_Serial_Array) ; unit
{
	editfield =
	For index, Element in  Combined_Serial_Array
	{
		;~ MsgBox % Combined_Serial_Array[A_Index]
		Editfield := Editfield element ",`n"
	}

	return Editfield
}

Formatted_Text_Serial_Count(Formatted_Text) ; unit
{
	Serial_Counter = 0
	Loop, Parse, Formatted_Text,`,
	{
		If (A_LoopField = "") || (A_LoopField = "`r") || (A_LoopField = "`n")
		continue
		Serial_Counter++
	}
	return Serial_Counter
}

Copy_selected_Text() ;Unit
{
	Clipboard =
	Send ^c ; sends a control C to the computer to copy selected text
	sleep(2)
	if clipboard =  ; if no text is seleted then clipboard will be remain blank
	return "No_Text_Selected"
	else
		return Clipboard ","
}




Check_For_Single_Serials(PreFormatted_Text) ; unit
{
	Loop, Parse, PreFormatted_Text, `,  ; loop to divide the Formatted_Text variable by the carraige returns
	{
		Debug_Log_Event("Check_For_Single_Serials() Loopfield is " A_LoopField)
		Prefix_Extract := Extract_Prefix(A_LoopField)
		Debug_Log_Event("Check_for_single_Serials() Prefix_Extract is " Prefix_Extract)

		If (Prefix_Extract = "`," or Prefix_Extract ="" or Prefix_Extract = "`n" or Prefix_Extract = "`r") ; checks to see if the Prefix_Store variable is a comma
		{
			Debug_Log_Event("Check_For_Single_Serials() continue")

			;msgbox, nothing there
			Prefix_Extract = ; sets the Prefix_Store variable to nothing
			Second_Number_set =  ; sets the Second_Number_Set variable to nothirng
			Continue ; skips over the rest of the loop and starts at the top of the parse loop
		}
If A_LoopField contains `-
{
Revised_PreFormatted_Text = %Revised_PreFormatted_Text%%A_LoopField%`,
Debug_Log_Event("Check_For_Single_Serials() Revised_PreFormatted_Text is " Revised_PreFormatted_Text)
continue
}
	StringTrimLeft, First_Number_Set,A_LoopField,3
	If First_Number_Set =
		Combine_Serials_together := Prefix_Extract

	else
		Combine_Serials_together := Prefix_Extract First_Number_Set "-" First_Number_Set

		Revised_PreFormatted_Text = %Revised_PreFormatted_Text%%Combine_Serials_together%`,


		Debug_Log_Event("Check_For_Single_Serials() First_Number_Set is " First_Number_Set)
		Debug_Log_Event("Check_For_Single_Serials() Revised_PreFormatted_Text is " Revised_PreFormatted_Text)
		Debug_Log_Event("Check_For_Single_Serials() New Parse" )
	}
	return Revised_PreFormatted_Text
}



Combinecount(Prefix_Store_Array) ; unit
{
	Prefixcombinecount = 0 ; Sets the Prefixcombinecount variable to 0

	Loop, Parse, Prefix_Store_Array, `,  ; parse loop to breaks the Prefixmatching variable up at the commas
	{
		if a_loopfield =  ; If the text before the comma is nothing, then skip the rest of the loop.
		Continue ; skip over the rest of the loop

		Prefixcombinecount++ ; Add one to Prefixcombinecount variable
	}
	return Prefixcombinecount
}

One_Up_All(Serial_Store_Array) ; unit
{
	One_Up_Prefix_array := Object()
	Length:= Serial_Store_Array.length()
	For index, Element in Serial_Store_Array
	{
		If A_index > %Length%
		Break
		Result := Extract_Prefix(Element)

		if (Element = "" or Element = "`," or Element = "`r" or Element = "`n")  ; If the text before the comma is nothing, then skip the rest of the loop.
		Continue ; skip over the rest of  the loop

		else
			One_Up_Prefix_array.insert(Result "00001-99999") ; Sets the One_Up_Prefix_array variable to the Prefix variable and adds in 00001-99999
	}

	return One_Up_Prefix_array
}



Extract_Prefix(Serial_Number) ; unit
{
	StringMid, Prefix_Store, Serial_Number, 1, 3 ; Gets the prefix letters and stores them for later reuse
	return Prefix_Store
}

Extract_First_Set_Of_Serial_Number(Serial_Number) ; unit
{
	Stringmid, First_Half_Serial_Num,Serial_Number,4,5 ; takes the numbers after teh prefix and stores them into String%Counting%number variable
	return First_Half_Serial_Num
}

Extract_Serial_Dividing_Char(Serial_Number) ; unit
{
	Stringmid, Check_Char,Serial_Number,9,1 ; takes the next char after the first half of the serial numbers
	return Check_Char
}

Extract_Second_Set_Of_Serial_Number(Serial_Number) ; unit
{
	Stringmid, Second_Half_Serial_Num,Serial_Number,10,5
	return Second_Half_Serial_Num
}



Checkvalues(Prefix_Store, First_Number_Set,  Second_Number_Set, Reset := 0) ; unit
{
	static Serial_Combine_Array := Object()



	oldprefix :=""


	If (Reset) ; This is here for the sole purpose of Unit testing. Since the Serial_Combine_array is static, is keeps the values between function call.The below will reset  the Serial_Combine_array to nothing.
	{
		Serial_array_reset := Object()

		Serial_Combine_Array := Serial_array_reset
		return
	}

	For Serial_index, Serial_element in Serial_Combine_Array
	{
		If  Serial_element contains %Prefix_store%
		{
			Prefix_Beg :=  Extract_First_Set_Of_Serial_Number(Serial_element)
			Prefix_Last :=  Extract_Second_Set_Of_Serial_Number(Serial_element)

			If 	Prefix_Beg > %Second_Number_Set%
			{
				Second_Number_Set = %Prefix_Beg%
			}

			If 	Prefix_Beg < %First_Number_Set%
			{
				First_Number_Set = %Prefix_Beg%
			}

			If 	Prefix_Last > %Second_Number_Set%
			{
				Second_Number_Set = %Prefix_Last%
			}

			Serial_Combine_Array[Serial_index] :=  Prefix_store First_Number_Set "-" Second_Number_Set
			return  Serial_Combine_Array
		}}

		Serial_Combine_Array.Insert(Prefix_store First_Number_Set . "-" . Second_Number_Set)

		return  Serial_Combine_Array
	}




	Remove_Formatting(Selected_Text) ; unit
	{
		StringReplace, Selected_Text,Selected_Text,`n,,All
		StringReplace, Selected_Text,Selected_Text,`r,,All
		StringReplace, Selected_Text,Selected_Text,`;,`,,All
		StringReplace, Selected_Text,Selected_Text,%A_Space%,,All
		StringReplace, Selected_Text,Selected_Text, `),`)`n,All
		StringReplace, Selected_Text,Selected_Text, 1`-up,,All
		StringReplace, Selected_Text,Selected_Text,  `-up,`-99999, All
		StringReplace, Selected_Text,Selected_Text, and,,All

		Debug_Log_Event("Remove_formatting()  return value is "  Selected_Text)

		Return Selected_Text
	}


	PreFormat_Text(Format_Removed_Text) ; unit
	{
		;Loops Format_Removed_Text variable to clean up the all the entereed text. Removes carriage returns with parse, removes any spaces, changes the ) to double ,
		; Changes 1-up to nothing, changes -up to 999999, changes ), to nothing
		Loop, Parse, Format_Removed_Text,`r`n`,
		{
			StringGetPos, pos, A_loopfield, :, 1
			Format_Text := SubStr(A_LoopField, pos+2)

			StringReplace, New_Format_Text,Format_Text,%A_Space%,,All
			StringReplace, New_Format_Text,New_Format_Text, `),`,,All
			StringReplace, New_Format_Text,New_Format_Text, 1`-up,,All
			StringReplace, New_Format_Text,New_Format_Text, `-up,`-99999, All
			StringReplace, New_Format_Text,New_Format_Text, `,,`,`n, All

			If (New_Format_Text != "") && (New_Format_Text !=" ,")
			Full_Text =  %Full_Text%%New_Format_Text%`,


			Debug_Log_Event("Preformat_Text()  Format_Text is "  Format_Text)
			Debug_Log_Event("Preformat_Text()  Format_Text is "  Format_Text)
			Debug_Log_Event("Preformat_Text()  New_Format_Text is "  New_Format_Text)
		}
			Return Full_Text
	}

	Prefix_Alone_Check_And_Add_One_UP(Text) ; unit
	{
;~ MsgBox, % "Test is `n" text
		Loop, parse, Text, `,
		{
			If (a_loopfield = "") || (A_LoopField = "`n") || (A_LoopField = "`r")
			{
				Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  A_LoopField is "  A_LoopField "`n Continue")
				Continue
			}
			Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  A_LoopField is "  A_LoopField )
			add_comma = %A_LoopField%`,
			StringGetPos, pos, add_comma, `,, 1

			If pos = 3
			{
				StringReplace, Added_Serial_Numbers,add_comma,`,,00001`-99999`,`n,all
			TextStore = %TextStore%%Added_Serial_Numbers%
			}
			Else if pos != 3
			{
				;~ Prefix:= Extract_Prefix(A_loopfield)
				;~ StringTrimLeft, Serial_Numbers, A_loopfield, 3
				;~ Serial_Numbers := add_digits(Serial_Numbers)
				;~ Serial_Numbers = %Prefix%%Serial_Numbers%
				;~ StringReplace , Serial_Numbers,Serial_Numbers,`),`,`n,all
				;~ TextStore = %TextStore%%Prefix%%Serial_Numbers%
				TextStore = %TextStore%%A_LoopField%`,`n
				}

		}

		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  Initial text is "  Text )
		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  Addcomma is "  Addcomma )
		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  pos is "  pos )
		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  Prefixstore is "  Prefixstore )
		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  Serial_Numbers is "  Serial_Numbers )
		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  StringTemp_end is "  StringTemp_end )
		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  Stringtemp is "  Stringtemp )
		Debug_Log_Event("Prefix_Alone_Check_And_Add_One_UP()  Textstore is "  Textstore )

		return Textstore
	}


	add_digits(Serial_Number) ; unit
	{
;~ MsgBox, % "Serial is "Serial_Number
			Final_Combined_Digits =
		combinestring =
		Loop, Parse, Serial_Number,`,`n
		{
			;~ MsgBox, % "Loopfield is  " A_LoopField
			If A_LoopField =
				continue

			Prefix  := Extract_Prefix(A_LoopField)
			StringReplace, Remove_End_Parenthesis,A_LoopField,`),,
			StringTrimLeft, Remove_End_Parenthesis,Remove_End_Parenthesis,3
			;~ MsgBox % "Remove end is " Remove_End_Parenthesis
			StringSplit, Serial_Temp, Remove_End_Parenthesis, `-
		Loop, 2
		{
		index := A_Index
   Loop % 5 - StrLen(Serial_Temp%A_index%)
       Serial_Temp%index% :="0"  Serial_Temp%index%
		}

		;~ combinestring := SubStr(combinestring, 2)

		Final_Combined_Digits =%Final_Combined_Digits%%Prefix%%Serial_temp1%`-%Serial_temp2%`,`n
		;~ MsgBox, % "final is  " Final_Combined_Digits
	}

		return Final_Combined_Digits
	}


	Clear_Format_Variables(unit_test) ; unit no need
	{
		global
		;Clear the variables
		Clipboard =
		Editfield =
		If (!Unit_Test)
		Guicontrol,, Editfield, %Editfield%
		return
	}






Export_Serials:
{
if checked = 0
Return

Progress, b w200 ,,Creating CSV file (This may take a minute or two)
Progress, 0

Effectivitycount = 1
WorkbookPath := A_Desktop "\Effectivity.CSV"    ; full path to your Workbook
Loop
{
IfNotExist %WorkbookPath%
Break

IfExist %WorkbookPath%
WorkbookPath := A_Desktop "\Effectivity" Effectivitycount ".CSV"    ; full path to your Workbook
Effectivitycount++
}


FileAppend, Prefix`,Begin Serial`,End Serial`n,  %WorkbookPath%


;msgbox, start export
GuiControlGet, EditField


Varcount = 0
Loop, parse, Editfield, `n
{
Varcount++
}

Percent_complete := 100 / Varcount


Loop, parse, Editfield, `n
{
	If A_LoopField contains - - - - - -
		continue

	Progress_total := Percent_Complete * A_Index
	Progress,  %Progress_total%
Prefix := Extract_Prefix(A_LoopField)
First_Number_Set := Extract_First_Set_Of_Serial_Number(A_LoopField)
Second_Number_Set := Extract_Second_Set_Of_Serial_Number(A_LoopField)
FileAppend, %Prefix%`,%First_Number_Set%`,%Second_Number_Set%`n,  %WorkbookPath%
}
Progress, off
Return
}
/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
ENter Serias Section
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/
Enter_Serials_Variable_Setup() ; no unit testing needed
{
	global
   Prefixcount = 5
     StartTime := A_TickCount
 }

Start_Macro() ; no unit testing needed as all contained functions are tested
{
	global
	GuiControlGet,Editfield
If Editfield =
{
Move_Message_Box("0","We Got A Problem","Oops! `nThere is no effectivity to input.`n`nPlease select the effectivity and press Ctrl +1. `nThank You!")
Exit
}
Gui_Image_Show("Run") ; options are Stop, Run, Pause, Start
 Enter_Serials_Variable_Setup()
Move_Message_Box("262144","Select ACM Screen","Click on the ACM window that you want to add effectivity to and then press the OK button`n`nNote that is the window is not full screen, the macro will make it full screen to increase macro reliability. ")
sleep(5)
WinGet, Active_ID, Id, A
SerialFullScreen(Active_ID)
CoordMode, mouse, Screen
Get_Add_Button_Screen_Position(Add_Button_X_Location, Add_Button_Y_Location)
   WinGet, Active_ID, Id, A
Get_Prefix_Button_Screen_Position(prefixx, prefixy)
Get_Apply_Button_Screen_Position(Applyx, Applyy)
Enter_Effectivity_Loop()
return
}
   ;~ gosub, Createtab

   ;~ Comma_Check(Effectivity_Macro)

    ;~ Tabcount = 0
   ;~ runcount = 21

Enter_Effectivity_Loop()
{
	
global breakloop, serialsentered, Applyx, Applyy, Effectivity_Macro, Refreshrate
LoopCount=0
 Loop
   {
      ;~ tabcount++
Load_ini_file(Configuration_File_Location)
checkforactivity()

;~ If Runcount > 20
	  ;~ {
	  ;~ runcount = 1
	  ;~ Gosub, GetSerial
	  ;~ Gosub, Copy_SerialTester
	  ;~ }
	  ;~ Send {ctrl down} %Tabcount% {Ctrl up}
	  ;~ if tabcount = 2
	  ;~ tabcount = 0


;~ sleep(10)

      If breakloop = 1
      {
         Gui 1: -AlwaysOnTop
         SplashTextOn,,,Macro Stopped
         Gui_Image_Show("Stop") ; options are Stop, Run, Pause, Start
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
					 Result :=   Searchend()
					  If (Result = Failure) or (Result = Timedout)
						Exit
						
					  Needrefresh = 0
					 ;msgbox, refresh
					 sleep(3)

					Win_check(Active_ID)
					 sleep()
					 Send {F5}
					 sleep(20)
						 Searchcountser = 0
								Loop
								{									
									If (breakloop)
										break
									
									 Click, %Applyx%,%Applyy%
									Click, %Applyx%,%Applyy%
										 
								Result :=   Searchend()
										  If (Result = Failure) or (Result = Timedout)
											Exit
								} until Result = Found

					 Loopcount = 0
					 Searchcountser = 0
					 Needrefresh = 0
					 sleep(10)
				  } ; End Refresh loop
	  
	  
	  
      sleep()
	Serial_number := Get_Serial_Numbers()
	if Serial_Number  contains  - - - - -
		{
		Stop_Issue_checks = 1
		Serial_number := Get_Serial_Numbers()
	}
     Prefix := Extract_Prefix(Serial_Number)

	 First_Effectivity_Numbers := Extract_First_Set_Of_Serial_Number(Serial_Number)
	 Second_Effectivity_Numbers := Extract_Second_Set_Of_Serial_Number(Serial_Number)

     If Prefix =
   {
      Complete = 1
    }
	
	Result := Searchend()
	If Result = Found
	      Enterserials(Prefix,First_Effectivity_Numbers,Second_Effectivity_Numbers, Active_ID, Complete)

If (Stop_Issue_checks)
{
Modifier = **Multiple Engineering Models**
Create_Dual_Instructions_GUI()
Counter = 0
loop
{
	Counter++
	if breakloop = 1
		Break
Result := Searchend()

Result_check := Searchend_Isssue_Check()
If (Result_check =" Issues_found") or (Result_check = "Dual_Eng")
	Result = no
If (Result = "Not_found") && (Result_check = "Not_found")
{

	If Counter = 10
	{
		Send {Enter 2}
	;~ Double_Click(Applyx,Applyy)
	Counter = 0
}

Gui, 70:Destroy
break
}
Gui 70: Flash
Sleep(10)
}
until (Result = "Found")

Loop
{
	if (breakloop)
		Break

;~ Double_Click(Applyx,Applyy)
Result := Searchend()
If  Result = Found
{
	Add_To_Completed_LIst(Serial_number, Modifier)
	break
}
}}

if (!Stop_Issue_checks)
{
	Loop
{
	If (breakloop)
		break
									
Sleep()
Loop, 2
{
	If (breakloop)
	break							
					Sleep(3)				
	Result := Searchend()
		If Result = Found
			{
			Modifier =
			Break
		}
}

Result := Check_For_Effectivity_Issues_Loop(Prefix,First_Effectivity_Numbers,Second_Effectivity_Numbers)

Send {Enter 2}
} until Result !=  Not_Found

If (Result  != "Bad Prefix") and  (Result != "Dual_Eng")
Add_To_Completed_LIst(Serial_number)
}
Serial_count := Added_Serial_Count()
		GuiControl,1:,serialsentered,%Serial_count%
		Gui,1:Submit,NoHide
   }
   return
   }

	Double_Click(x,y) ; no unit test needed
	{
Click %x%, %y%
Click %x%, %y%
}

Check_For_Effectivity_Issues_Loop(Prefix,First_Effectivity_Numbers,Second_Effectivity_Numbers)
{	
	global breakloop
Loop, 8
{
	If (breakloop)
		break

Result :=   Searchend_Isssue_Check()
			If (Result = "Dual_Eng")
			{
				activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
				SplashTextOn,,,Dual Serial or Eng Models, This serial has dual Engineering Models. This Prefix will be moved to the End of the list.
				Winmove, ,This serial has dual engeering Models,%amonx%, %Amony%
				Multiple_Eng_Model_Move_To_End(Prefix,First_Effectivity_Numbers,Second_Effectivity_Numbers)
				Added_Serial_Count("-1")
				Return Result
				Break
		  }
			if (Result = "Bad Prefix")
			{
			;~ MsgBox, Bad Prefix
			Serialnogo(Prefix,First_Effectivity_Numbers,Second_Effectivity_Numbers)
			Added_Serial_Count("-1")
			   Return Result
			   Break
	  }
	  	
;~ Double_Click(Applyx,Applyy)
	}

   Return "Not_Found"
}


	Create_Dual_Instructions_GUI() ; no unit test needed
	{
	activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
Gui 70:Add,Text,, Select an engeering model
Gui  70:Show, x%amonx% y%amony%,%Effectivity_Macro%
gui, 70: +AlwaysOnTop
Gui 70: Flash
return
}
Added_Serial_Count(Add_Or_Subtract := "1") ; unit
{
	static Add_Count
	Add_count += %Add_Or_Subtract%
return Add_count
}


Enterserials(Prefix_Holder_for_ACM_Input,First_Effectivity_Numbers,Second_Effectivity_Numbers, Active_ID, Complete)
{

	global  prefixx, prefixy, Applyx, Applyy, Add_Button_X_Location, Add_Button_Y_Location, StartTime, Sleep_Delay

If (!Complete)
{
CoordMode, mouse, Screen
   ;listlines on
    win_check(Active_ID)
   Click, %prefixx%, %prefixy%
   sleep()
   mousemove 300,300
   sleep()
   SEndRaw, %Prefix_Holder_for_ACM_Input%
   sleep()
   Send {Tab}
 win_check(Active_ID)
   Sendraw, %First_Effectivity_Numbers%
   sleep(3)
   Send {Tab}
  win_check(Active_ID)
   SendRaw, %Second_Effectivity_Numbers%
   sleep(3)
   Send {Tab}
   Sleep(Sleep_Delay)
   Send {enter 2}
   Sleep(2)
Double_Click(Applyx,Applyy)
Sleep(3)
   Serialcount +=1
   sleep()
   Searchcount = 0
   Searchcountser = 0
}
else
	  {
      ElapsedTime := A_TickCount - StartTime
      Total_Time := milli2hms(ElapsedTime, h, m, s)
      sleep(5)
      Send {f5}
       Move_Message_Box("0",""," The number of successful Serial additions to ACM is "  Serialcount "`n`nMacro Finished due to no more Serials to add. `n`n It took the macro " Total_Time " to perform tasks. `n`n Please close Serial Macro Window after checking to ensure serials were entered correctly.")
      Guicontrol,1:, Editfield,
      ;~ gosub, radio2h
  Gui_Image_Show("Stop")
      Exit
   }
   Return
}


/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
Macro Timeout Functions and GUI
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/


Serialnogo(Prefix,First_Effectivity_Numbers,Second_Effectivity_Numbers)
{
	;~ MsgBox, Serialnogo prefix is %Prefix%
	Serial_Number := Prefix First_Effectivity_Numbers "-" Second_Effectivity_Numbers
	Modifier = ** Not In ACM **
	Add_To_Completed_LIst(Serial_number,Modifier)
	GuiControlGet, Editfield
Loop, Parse, Editfield, `n
{
Prefix_Check := Extract_Prefix(A_LoopField)
If Prefix_Check = %Prefix%
{
Add_To_Completed_LIst(A_LoopField, Modifier)
continue
}
else
EditfieldStore := EditfieldStore A_LoopField "`n"
}
Guicontrol,,Editfield, %EditfieldStore%
Clear_ACM_Fields()
return
}

Add_To_Completed_LIst(Serial_number, Modifier := "")
{
	GuiControlGet, Editfield2
	GuiControl,, Editfield2, %Editfield2%`n%Serial_Number%  %Modifier%
Return
}

Clear_ACM_Fields()
{
	global Editfield2, Active_ID, prefixx, prefixy

win_check(Active_ID)
   Click, %prefixx%, %prefixy%
   sleep(5)
   Send {ctrl down}{a}{Ctrl up}
   sleep()
   Send {Del}{Tab}
   Send {Del}{Tab}
   Send {Del}{Tab}
   Click, %prefixx%, %prefixy%
   Send {Del}
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


Multiple_Eng_Model_Move_To_End(Prefix,First_Effectivity_Numbers,Second_Effectivity_Numbers)
{
Afterloop := Prefix First_Effectivity_Numbers "-" Second_Effectivity_Numbers
GuiControlGet, Editfield
Loop, Parse, Editfield, `n
{
	If (A_LoopField = "`n") or (A_LoopField = "")
		continue

	If A_LoopField contains  - - - -
	{
		Stop_Check = 1
			EditfieldStore := EditfieldStore A_LoopField
			continue
	}

If (!Stop_check)
{
Prefix_Check := Extract_Prefix(A_LoopField)
If Prefix_Check = %Prefix%
Afterloop := Afterloop "`n" A_LoopField
else
EditfieldStore := EditfieldStore A_LoopField "`n"
}
else
	EditfieldStore := EditfieldStore A_LoopField "`n"
}
Guicontrol,,Editfield, %EditfieldStore%`n%Afterloop%
Clear_ACM_Fields()
return
}



/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
Supporting Functions
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/
milli2hms(milli, ByRef hours=0, ByRef mins=0, ByRef secs=0, secPercision=0) ; no unit testing needed
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

Win_check(Active_ID) ; no unit test needed
{
WinGetTitle, Title, ahk_id %Active_ID%
   IfWinNotActive , %Title%
   {
      WinActivate, %Title%
      WinWaitActive, %Title%,,3
      sleep(5)
   }

   return
}


checkforactivity() ; no unit test needed
{
	global breakloop, Active_ID
   while A_TimeIdlePhysical < 4999 ; meaning there has been user activity
   {
      Gui 1: -AlwaysOnTop

      {
         If breakloop = 1
         {
            SplashTextOff
			Break
            Return
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
}
         SplashTextOn,350,50,Macro paused, Macro is now paused due to user activity.`n Macro will resume after %timeleft% seconds of no user input
		        Gui_Image_Show("Pause") ; options are Stop, Run, Pause, Start
         Gui, Submit, NoHide
         sleep(10)
      }

      if A_TimeIdlePhysical > 5000 ; meaning there has been no user activity
      {
           splashtextoff
            Gui 1: +AlwaysOnTop
			win_check(Active_ID)
		       Gui_Image_Show("Run") ; options are Stop, Run, Pause, Start
            ;~ gosub, radio1h
            Gui, Submit, NoHide
            ;~ gosub, SerialFullScreen
         }
      return
   }


 Get_Serial_Numbers()
 {
	global Effectivity_Macro, nextserialtoadd
	 COntrolSEnd,,{Ctrl Down}{Home}{Ctrl Up}, %Effectivity_Macro% ;note that the controlsend has two commas after the function call (THis always messed me up)
   sleep()
   COntrolSEnd,,{Shift Down}{End}{Shift Up}, %Effectivity_Macro%
   Sleep()
   ControlGet,Serial_Number,Selected,,,%Effectivity_Macro%

   COntrolSEnd,,{Del 2}, %Effectivity_Macro%
   GuiControl,,nextserialtoadd,%Serial_Number%
return Serial_Number
}


   SerialFullScreen(Active_ID) ; no unit test needed
   {
      WinGetPos, Xarbor,yarbor,warbor,harbor, ahk_id %Active_ID%
      CurrmonAM := GetCurrentMonitor()
      SysGet,Aarea,MonitorWorkArea,%CurrmonAM%
      WidthA := AareaRight- AareaLeft
      HeightA := aareaBottom - aAreaTop
      lefta := (aAreaLeft - 4) ; this is -4 becuase Oracle puts a 4 pixel border on its seamless windows
      topa := (AAreaTop - 4) ; this is -4 becuase Oracle puts a 4 pixel border on its seamless windows
;~ MsgBox % WidthA HeightA
	  ;~ MsgBox, % "Left is "leftt " top is " topp "`nxarbor is " xarbor " yarbor is " yarbor
      MouseGetPos mmx,mmy
      If (yarbor = topa)
      {
         If (xarbor = lefta)
         ;Msgbox, win maxed
         Return
      }
      Else
         ;msgbox, not maxed
      CoordMode, mouse, Relative
      MouseMove 300,10
      Click 2
      ;~ Click
      Coordmode, mouse, screen
      ;MouseMove, mmx, mmy
      return
   }

GetCurrentMonitor() ; no unit test needed
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



Searchend()
{
global Image_Red_Exclamation_Point, Active_ID, 
   ;msgbox, searchend
   listlines off

	Current_Monitor := GetCurrentMonitor()
	pToken := Gdip_Startup()
	bmpNeedle1 := Gdip_CreateBitmapFromFile(Image_Red_Exclamation_Point)
   ;~ bmpHaystack := Gdip_BitmapFromScreen(Current_Monitor)
   bmpHaystack :=    Gdip_BitmapFromHWND(Active_ID)
   RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,0,0,0,0)
   Gdip_Shutdown(pToken)
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
  Move_Message_Box("262144", Effectivity_Macro, "Error Searchend (bmpNeedle1)" RETSearch)
      Exit
   }


   If RETSearch > 0
   {
      ;~ SetTimer, refreshcheck, Off
      Refreshchecks = 0
      SleepStill = 0
      ;Msgbox, found
	  Return "Found"
	}
	else
   Return "Not_Found"
}


Searchend_Isssue_Check()
{
global Issues_Image, Active_ID, prefixx, prefixy

   listlines off
	Current_Monitor := GetCurrentMonitor()
	pToken := Gdip_Startup()
	bmpNeedle1 := Gdip_CreateBitmapFromFile(Issues_Image)
   bmpHaystack :=    Gdip_BitmapFromHWND(Active_ID)

   RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedle1,,0,0,0,0,0,0,0,0)

   Gdip_Shutdown(pToken)
      ;listlines on
	  ;~ MsgBox, % RETSearch
   If RETSearch < 0
   {
      if RETSearch = -1001
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
  Move_Message_Box("262144", Effectivity_Macro, "Error Searchend Issues check  (bmpNeedle1)" RETSearch)
      Exit
   }

;~ MsgBox, % RETSearch

   If (RETSearch = "4") or (RETSearch = "2")
   {
	  Return "Dual_Eng"
   }
    If (RETSearch = "3") or (RETSearch = "5") or (RETSearch = "6")
   {
	  Return "Bad Prefix"
	}
		 If (RETSearch = "0")
   Return "Not_found"

		 If (RETSearch != "0") or (RETSearch != "4") or (RETSearch != "5") or (RETSearch != "6")
   Return "Issues_found"
}
/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
Below is the Functions to support the initial starting before entering the seirals
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/

Comma_Check(Effectivity_Macro)
{
   COntrolSEnd,, {Shift Down}{Right}{SHift Up},%Effectivity_Macro%
   ControlGet,Commacheck,Selected,,,%Effectivity_Macro%
   If Commacheck = ,
   {
      ;Msgbox, COmmafound
   }else  {
      ;msgbox, comma not found
      COntrolSEnd,, {Left}{BackSpace 2}, %Effectivity_Macro%
   }
   Return
}


Wait_For_Shift_Mouse_Click() ; no unit test needed
{

   Keywait, Shift,D
   Keywait, Lbutton, D
   KeyWait, Shift
	return "Done"
}


Get_Add_Button_Screen_Position(ByRef X_Location, ByRef Y_Location) ; no unit test needed
{
   SetTimer, ToolTipTimerbutton, 10  ;timer routine will occur every 10ms..

Wait_For_Shift_Mouse_Click()

   MouseGetPos, X_Location, Y_Location
sleep()
   SetTimer, ToolTipTimerbutton,Off
   Tooltip,
   return
}

ToolTipTimerbutton:
{
   tooltip, Please Shift + mouse button click on the "Add Button" in the ACM effectivity screen to get its position.
   Return
}


Get_Prefix_Button_Screen_Position(ByRef X_Location, ByRef Y_Location) ; no unit test needed
{
   settimer, ToolTipTimerprefix,10

Wait_For_Shift_Mouse_Click()

   MouseGetPos, X_Location, Y_Location
sleep()
    settimer, ToolTipTimerprefix,Off
   Tooltip,
   return
}

ToolTipTimerprefix:
{
ToolTip, Please Shift + mouse button click in the "prefix" edit field in the ACM effectivity screen to get it's location.
   Return
}


Get_Apply_Button_Screen_Position(ByRef X_Location, ByRef Y_Location) ; no unit test needed
{

   SetTimer, ToolTipTimerapply, 10  ;timer routine will occur every 10ms..

Wait_For_Shift_Mouse_Click()

   MouseGetPos, X_Location, Y_Location
sleep()
   SetTimer, ToolTipTimerapply,Off
   Textapplybutton =
   Tooltip,
   return
}


ToolTipTimerapply:
{
ToolTip, Please Shift + mouse button click on the "Apply button" in the ACM effectivity screen to get it's location.
   Return
}

/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Below is the Create_Main__gui_Menu funciton along with the functions it solely calls to .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

	Create_Main_GUI_Menu() ; no unit test needed
	{
		Menu, BBBB, Add, &Check For Update , Versioncheck
		Menu, BBBB, Add, &Options, OptionsGui
		Menu, BBBB, Add,
		Menu, CCCC, Add, &Run							(Crtl + 2), Start_Macro
		Menu, CCCC, Add, &Pause/Unpause 				(Pause / Insert), pausesub
		Menu, CCCC, Add, &Stop Macro					(ESC), Exit_Program
		Menu, CCCC, Add, &Reload Macro, restart_macro
		Menu, CCCC, Add, &Reload Macro with Current Effectivity, restart_macro_Effectivity
		Menu, BBBB, Add, &Exit							(Ctrl + Q), Quitapp
		Menu, DDDD, Add, &How To Use					(F1), HowTo
		Menu, DDDD, Add, &About , Aboutmacro

		Menu, MyMenuBar, Add, &File, :BBBB
		Menu, MyMenuBar, Add, &Macro, :CCCC
		Menu, MyMenuBar, Add, &Help, :DDDD
		Return
	}

	pausesub:
	{
		if A_IsPaused = 0
		{
			gosub, radio_button
			Gui 1: -AlwaysOnTop
			Gui_Image_Show("Pause") ; Options are Start, Pause, Run, Stop
			Gui, Submit, NoHide
			Loop, 4
			{
				Move_Message_Box("262144",Effectivity_Macro, "Macro is paused. Press pause key to unpause",".1")
			}
			Move_Message_Box("262144",Effectivity_Macro, "Macro is paused. Press pause key to unpause","10")
			Pause, toggle, 1
			Return
		}else  {
			Gui 1: +AlwaysOnTop
			Gui_Image_Show("Run") ; Options are Start, Pause, Run, Stop
			Gui, Submit, NoHide
			Pause, toggle, 1
		}
		return
	}

	Exit_Program(Unit_Test := 0) ; unit
	{
		global Serialcount
		Result := Move_Message_Box("262148",Effectivity_Macro, " The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure you want to Quit the macro?.`n`n Press YES to Quit the Macro.`n`n No to keep going.")

		If Result = Yes
		{
			Stopactcheck = 1
			Gui 1: -AlwaysOnTop
            If (!unit_test)
			Gui_Image_Show("Stopped")

			Send {Shift Up}{Ctrl Up}
			breakloop = 1
            If (!Unit_Test)
			ExitApp
		}
		return Result
	}

Clickyes:
{
Send {y}
SetTimer, Clickyes, Off
return
}

Clickno:
{
Send {n}
SetTimer, Clickno, Off
return
}

	restart_macro() ; no unit test needed
	{
		Result := Move_Message_Box("262148",Effectivity_Macro, "Are you sure that you want to reload the program?" )
		If Result =  yes
		Reload

		return
	}

	restart_macro_Effectivity() ; no unit test needed
	{
		global nextserialtoadd, EditField, EditField2,reloadprefixtext,serialsentered

		Result := Move_Message_Box("262148",Effectivity_Macro, "Are you sure that you want to reload the program?" )

		If Result =  yes
		{
			GuiControlGet, nextserialtoadd
			GuiControlGet, EditField
			GuiControlGet, EditField2
			GuiControlGet, Serialcount
			GuiControlGet, reloadprefixtext

			If nextserialtoadd =
			EditField = %editfield%
			else
				EditField = %nextserialtoadd%`n%editfield%

			FileAppend, %EditField%, C:\SerialMacro\TempAdd.txt
			FileAppend, %EditField2%, C:\SerialMacro\TempAdded.txt
			FileAppend, %reloadprefixtext%, C:\SerialMacro\Tempamount.txt
			FileAppend, %serialsentered%, C:\SerialMacro\Tempcount.txt
			Reload
		}

		return
	}


	/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Gui Screens .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

	Serials_GUI_Screen(editfieldamount, editfield2amount, TotalPrefixestemp) ; no unit test needed
	{
		Global

		activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

		If TotalPrefixestemp < 1
		{
			TotalPrefixestemp = 0
		}
		gui 1:add, Edit, x10 y50 w390 h240  vEditField,%editfieldamount%
		gui 1:add, Edit, xp yp w390 h240 vEditField2,%editfield2amount%

		Gui 1:Add, Picture, x315 y310 w50 h50 +0x4000000  BackGroundTrans vStarting gstart_macro , C:\SerialMacro\images\Start.png
		Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vRunning, C:\SerialMacro\images\Running.png
		Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vpaused  gpausesub, C:\SerialMacro\images\Paused.png
		Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vStopped grestart_macro, C:\SerialMacro\images\Stopped.png
		Gui, 1:Add, Picture, x0 y0 w410 h400 +0x4000000 , C:\SerialMacro\images\background.png

		Gui 1:Add, Edit, xp+165 yp+343 w110 h20  vnextserialtoadd, %nextserialtoaddv%

		Gui 1:Add, Text, x5 y5   BackgroundTrans +Center , There are a total of
		Gui 1:Add, Text, xp+65 w75  BackgroundTrans +Center vreloadprefixtext, %TotalPrefixestemp%
		Gui 1:Add, Text, xp+55  BackgroundTrans +Center , Effectivity to add to ACM

		Gui 1:add, Radio, xp-97 yp+25 w130 h20 BackGroundTrans Checked vradio1 gradio_button, Effectivity to be added
		Gui 1:add, Radio, xp+155 yp w140 h20 BackGroundTrans  vradio2 gradio_button, Effectivity already added

		Gui 1:Add, Text, xp-170 Yp+265 W250 h13 BackGroundTrans, Number of Effectivity successfully added to ACM =
		Gui 1:Add, Text, xp+245  h13 BackGroundTrans vserialsentered, %Serialcount%

		Gui 1:Add, Text, Xp-245 Yp+15 w250 h13  BackGroundTrans , If macro is operating incorrectly, press Esc to reload
		Gui 1:Add, Text, xp yp+15 w250 h13  BackGroundTrans , Or press Pause Button on keyboard to Pause macro, Press Pause again to resume macro.
		Gui 1:Add, Text, xp yp+20 w145 h20  BackgroundTrans , Next Effectivity to add to ACM:
		Gui 1:Menu, MyMenuBar
		Gui 1:Show,  x%amonx% y%amony% , %Effectivity_Macro%
		gui 1: +alwaysontop
		Editfield_Control("Editfield")
		Guicontrol,,Radio1,1

		; IfExist  C:\SerialMacro\Tempcount.txt
		; {
		; FileRead, Serialcount,C:\SerialMacro\Tempcount.txt
		; GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%
		; FileDelete, C:\SerialMacro\Tempcount.txt
		; }
		Gui 1:Submit, NoHide
		Return
	}

	radio_button:
	{
		Gui,1:submit,nohide
		If (Radio1)
		{
			GuiControl,, Radio1, 1
			Editfield_Control("Editfield")
		}

		If (radio2)
		{
			Editfield_Control("Editfield2")
			GuiControl,, Radio2, 1
		}

		gui, 1:submit, nohide
		return
	}

	Editfield_Control(Textbox, Gui_Number := 1) ; unit
	{
      global editfield, editfield2

		If Textbox = Editfield
		{
			Guicontrol,%Gui_Number%: hide, Editfield2,
			Guicontrol,%Gui_Number%: show, Editfield,
			Guicontrol,%Gui_Number%: Focus, Editfield
		}else  {
			Guicontrol,%Gui_Number%: hide, Editfield,
			Guicontrol,%Gui_Number%: show, Editfield2,
			Guicontrol,%Gui_Number%: Focus, Editfield2
		}
		return
	}


	Gui_Image_Show(Image)
	{
		global paused, Starting, Stopped, Running
		If ( image = "Pause")
		{
			Guicontrol,show, paused
			Guicontrol,hide, Starting
			Guicontrol,hide, Stopped
			Guicontrol,hide, Running
		}

		If (image = "Run")
		{
			Guicontrol,Show, Running
			Guicontrol,Hide, paused
			Guicontrol,hide, Stopped
			Guicontrol,hide, Starting
		}
		If (image = "Stop")
		{
			Guicontrol,Show, Stopped
			Guicontrol,hide, Starting
			Guicontrol,hide, paused
			Guicontrol,hide, Running
		}
		If (image = "start")
		{
			Guicontrol,show, Starting
			Guicontrol,hide, paused
			Guicontrol,hide, Stopped
			Guicontrol,hide, Running
		}
		return
	}


	Temp_File_Read(File_Install_Root_Folder,File_Name) ; unit
	{
		IfExist, %File_Install_Root_Folder%\%File_Name%
		{
			FileRead, Variable_Store, %File_Install_Root_Folder%\%File_Name%

			Debug_Log_Event("Temp_File_read() ......File name is " File_Name " Variable_store is " Variable_Store)
		}
		else
			Variable_Store = Null

		return Variable_Store
	}

	Temp_File_Delete(File_Install_Root_Folder,File_Name) ; unit
	{
		IfExist, %File_Install_Root_Folder%\%File_Name%
		{
			FileDelete, %File_Install_Root_Folder%\%File_Name%

			Debug_Log_Event("Temp_File_Delete() ......File name is " File_Name)
			return  File_Name " - File Found and Deleted"
		}
		else
			return  File_Name " - File Not Exist"
	}

	Move_Message_Box(Msg_box_type,Msg_box_title, Msg_box_text, Msg_box_Time := 2147483 ) ; no unit test needed
	{
		global
		   Gui 1: -AlwaysOnTop
		activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
		Settimer, winmovemsgbox, 20
		MsgBox, % Msg_box_type , %Msg_box_title% , %Msg_box_text% , %Msg_box_time%
  Gui 1: +AlwaysOnTop
		IfMsgBox yes
		Result = Yes
		IfMsgBox no
		Result = no
		IfMsgBox Ok
		Result = Ok
		IfMsgBox  Cancel
		Result = Cancel
		IfMsgBox Abort
		Result = Abort
		IfMsgBox Ignore
		Result = Ignore
		IfMsgBox Retry
		Result = Retry
		IfMsgBox continue
		Result = continue
		IfMsgBox TryAgain
		Result = TryAgain
		IfMsgBox Timeout
		Result = Timeout

		return Result
	}

	winmovemsgbox:
	{
		SetTimer, WinMoveMsgBox, OFF
		WinMove, %Msg_box_text% , Amonx, Amony
		return
	}

	activeMonitorInfo( ByRef aX, ByRef aY, ByRef aWidth,  ByRef  aHeight, ByRef mouseX, ByRef mouseY  ) ; no unit test needed
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
				ax 		:= (aWidth / 2)
				ay		:= (aHeight / 2)
				;msgbox, ax is %ax% `n ay is %ay% `n aheight is %aHeight% `n awidth is %aWidth%
				return
			}}}

			SerialbreakquestionGUI() ; no unit test needed
			{
				global createexcel, checked
				activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

				gui 1: -alwaysontop
				Gui, 8:Add, Picture, x0 y0 w400 h90 +0x4000000, %File_Install_Work_Folder%\images\background.png
				Gui, 8: Add, text, x10 y20 BackgroundTrans, Do you want to combine the serial breaks, or keep the serial breaks seperated?
				Gui, 8:add, button, xp+50 yp+20 gcombinequstion, Combine
				gui, 8:add, button, xp+75 yp gkeepseperated, Keep Seperated
				gui, 8:add, button, xp+115 yp goneup, 1-UP all Effectivity
				Gui, 8:Add, Checkbox, XP-190 yp+30 vcreateexcel, Export Effectivity to Excel file (Effectivity.CSV)
				Gui, 8:show, x%amonx% y%amony% w400 h90
				gui 8: +alwaysontop
				Pausescript()			
				return checked
			}

			Pausescript() ; no unit test needed
			{
				Menu,Tray,Icon, % "C:\Serialmacro\icons\paused.ico", ,1
				Gui_Image_Show("Pause")
				Pause,on
				Return
			}

			UnPausescript() ; no unit test needed
			{
				Pause,off
				Menu,Tray,Icon, % "C:\Serialmacro\icons\Serial.ico", ,1
				Gui_Image_Show("Run")	
				Return
			}


			oneup()
			{
				global checked
				UnPausescript()
				GuiControlGet, checked,, createexcel
				gui 1: +alwaysontop
				Oneupserial = 1
				Gui, 8:destroy
			Return Checked
			}

			combinequstion()
			{
				global checked
				UnPausescript()
				GuiControlGet, checked,, createexcel
				gui 1: +alwaysontop
				Gui, 8:destroy
				combine = 1
				Return Checked
			}

			keepseperated()
			{
				global checked
				UnPausescript()
				GuiControlGet, checked,, createexcel
				gui 1: +alwaysontop
				Gui, 8:destroy
				combine = 0
				Oneupserial = 0
				Return Checked
			}

			aboutmacro() ; no unit test needed
			{
				global File_Install_Work_Folder, Program_Location_Link, Effectivity_Macro
				activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

				Gui, 7:Add, Picture, x0 y0 w525 h300 +0x4000000, %File_Install_Work_Folder%\images\background.png
				gui, 7:add, Text, xp+5 yp+5 w500 h40 BackgroundTrans, This program was designed to help increase the speed and accuracy of entering machine Effectivity prefixes form the Pubtool into ACM.
				gui, 7:add, Text, xp yp+45 w500 h20 BackgroundTrans, To get the latest version, go to the File menu and select Check for updates.
				gui, 7:add, Text, xp yp+25 w300 h20 BackgroundTrans,  The location of the macro is at the following box account:
				gui 7:font, CBlue Underline
				gui, 7:add, Text, xp+275 yp w500 h20 BackgroundTrans gboxlink, %Program_Location_Link%
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

			emaillink() ; no unit test needed
			{
				Run,  mailto:Karnia_Jarett_S@cat?Subject=Effectivity Macro
				return
			}

			boxlink() ; no unit test needed
			{
				global Program_Location_Link
				Run, %Program_Location_Link%
				return
			}


			sleep(Amount := 1) ; no unit test needed
			{
				ListLines off
				amount := amount * 100
				Sleep %Amount%
ListLines on
Return
			}


	/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Update macro functions  .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

			Versioncheck() ; no unit test needed
			{
				global Version_Number, Configuration_File_Location, Update_Check_URL, First_Run

				Load_ini_file(Configuration_File_Location)

IfExist, %A_desktop%\testdownload.txt
FileDelete, %A_desktop%\testdownload.txt
	
				Progress,  w200, Updating..., Gathering Information, Effectivity Macro Updater
				Progress, 0
				
					URLDownloadToFile,%Update_Check_URL%, %A_desktop%\TestDownload.txt

				Progress,  w200, Updating..., Fetching Server Information, Effectivity Macro Updater
				Progress, 15

				Progress,  w200, Updating...,Gathering Current Version From Server, Effectivity Macro Updater
				Progress, 50
		
				Progress,  w200,Updating..., Comparing Version Information, Effectivity Macro Updater
				Progress, 60
				
Loop, Read, %A_desktop%\testdownload.txt
{
	If A_LoopReadLine contains Version=
		update_Version := A_LoopReadLine
	else
		What_is_new_text := What_is_new_text A_LoopReadLine "`n"
}

StringReplace, update_Version,update_Version,Version=,,

				If (update_Version <= Version_Number) and  (First_run = "0")
				{
					Progress,  w200,Updating..., Macro is Up to date., Effectivity Macro Updater
					Progress, 100
					sleep(10)
				}

				If (update_Version > Version_Number)  or (First_run = "1")
				{
				Progress, Off
					gui,35: font, S15  ;Set 10-point Verdana.
				Gui, 35:Add, Text,x5 y5, New Version available!
				Gui, 35:Add, Text,xp yP+25, Your version is %Version_Number% 
			gui,35: font, S15 cRED  ; Set 10-point Verdana.
				Gui, 35:Add, Text,xp yP+25, New  version is %update_Version% 
					gui,35: font, s10 cblack  ; Set 10-point Verdana.
					If (First_Run)
				Gui, 35:Add, Edit,xp yP+35 w600 h500,  Looks Like This is Your first time Running This Version. `n`n %What_is_new_text%
				else
				Gui, 35:Add, Edit,xp yP+35 w600 h500, %What_is_new_text%
				If (First_run)
				Gui, 35:Add, Button, yp+525  w100 h25 gCancel, Ok
				else
				{
				Gui, 35:Add, Button, yp+525 gDownload_new_version, DOWNLOAD NEW VERSION
				Gui, 35:Add, Button, xp+200 gCancel, Cancel
			}
				gui, 35:Show,,New Version!
					First_Run=0
						Write_ini_file(Configuration_File_Location)
				Pause, on
				}
Progress, Off
				return
			}
			
35GuiEscape:
35GuiClose:
{
	gui, 35:destroy
Pause, Off
Return
}

Download_new_version()
{
	global Program_Location_Link
	Pause, off
	Gui 35: Destroy
Run, %Program_Location_Link%
	
	return
}

Cancel:
{
	pause, off
Gui 35:Destroy
return
}	

				OptionsGui:
				{
			     	activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
					Load_ini_file(Configuration_File_Location)
					;~ IniRead, Refreshrate, %Configuration_File_Location%, Refreshrate,Refreshrate
					;~ IniRead, Sleep_Delay, %Configuration_File_Location%, Sleep_Delay,Sleep_Delay
					gui 10: +alwaysontop
					Gui , 1: -AlwaysOnTop
					gui 10:add, text, x5 y5 w320 h20 ,Refreash ACM Rate (After how many entered effectivity)
					gui 10:add, edit, xp+275 yp-3 w30 veditfield5 , %refreshrate%
					gui 10:add, Text, xp-275 yp+30, ACM speed Compensation (10 = 1 second delay)
					gui 10:add, edit, xp+275 yp-3 w30 veditfield10 , %Sleep_Delay%
					gui 10:add, button, xp-251 yp+26 h20 w75 Default gsavesets, Save Settings
					Gui, 10:Add, Picture, x0 y0 w325 h100 +0x4000000 , %File_Install_Work_Folder%\images\background.png
					gui 10:show, x%amonx% y%amony% w325 h100, Options
					Guicontrol,10:, editfield5, %refreshrate%
					Guicontrol,10:, editfield10, %Sleep_Delay%
					gui, 10:submit, nohide
					return
				}

				GuiClose:
				{
					Result := Move_Message_Box("262148","Quit" Effectivity_Macro, "Are you sure you want to quit?")
					;~ MsgBox, %Program_Location_Link%
					If Result =  yes
					ExitApp

					return
				}

				10GuiClose:
				10Guiescape:
				{
					Gui, 10: Destroy
					return
				}

				savesets()
				{
					global
					Gui , 1: +AlwaysOnTop
					gui 10:submit, nohide
					GuiControlGet,Refreshrate,,editfield5
					GuiControlGet,sleep_delay,,editfield10
				IniWrite, %Sleep_Delay%, %Configuration_File_Location%,Sleep_Delay,Sleep_Delay
				IniWrite, %Refreshrate%, %Configuration_File_Location%,Refreshrate,Refreshrate
					gui 10:destroy
					return
				}

savevalue:
{
GuiControlGet,Sleep_Delay,,EditField20
IniWrite, %Sleep_Delay%, %Configuration_File_Location%,Sleep_Delay,Sleep_Delay
return
}







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

