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

;~ #include Unit_testing\Unit_testing.ahk  ; Uncomment this to run unit test modules, to narrow down what function is broken
/*
****************************************************************************************************************************************************
************ Variable Setup *******************************************************************************
*****************************************************************************************
*/

Global Prefix_Number_Location_Check, First_Effectivity_Numbers, Title, sleepstill, Current_Monitor, Log_Events, Unit_test, File_Install_Work_Folder, Oneupserial, combineser

Version_Number = 1.4 beta
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

Unit_Test = 1 ; Set this to 1 to perform unit tests and logging.
Log_Events = 0 ;Set this to 1 to perform logging


inifile = c:\Serialmacro\config.ini
File_Install_Root_Folder = C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files

File_Install_Work_Folder = C:\SerialMacro

Image_Red_Exclamation_Point = %File_Install_Work_Folder%\red_image.png
IMage_Actve_Add_Button = %File_Install_Work_Folder%\Active_plus.png
Image_Active_Apply_Button = %File_Install_Work_Folder%\orange_button.png



Result = Folder_Exist_Check("SerialMacro")
If Result contains Folder_Not_Exist
Folder_Create("SerialMacro")

Result := Folder_Exist_Check("SerialMacro\Images")
If Result contains  Folder_Not_Exist
Folder_Create("SerialMacro\Images")

Result := Folder_Exist_Check("SerialMacro\icons")
If Result contains Folder_Not_Exist
Folder_Create("SerialMacro\icons")

Result := File_Exist_Check("config.ini")
If Result  contains File_Not_Exist
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
	Load_ini_file(inifile)
}

If Refreshrate = Error
{
	IniWrite, 20,  %inifile%,refreshrate,refreshrate
	Load_ini_file(inifile)
}



;~ Result := Install_Requied_Files_Root(File_Install_Work_Folder)
;~ If (Result)
	;~ Move_Message_Box("0","Error","Error with installing How To PDF File. `n`nThis will not effect Program Operation")

;~ Result := Install_Requied_Files_Icons(File_Install_Work_Folder)
;~ If (Result)
	;~ Move_Message_Box("0","Error","Error with installing Program Icons. `n`nThis will not Effect Program Operation")

;~ Install_Requied_Files_Images(File_Install_Work_Folder)
;~ If (Result)
;~ {
	;~ Move_Message_Box("0","Error","Error with installing Images needed for screen searching.`n`n Program will not run without these files.`n`n Please restart Program and if error occurs agian, Contact Jarett Karnia for assisstance")
	;~ If (!unit_test)
		;~ ExitApp
;~ }

Create_Tray_Menu()
Create_Main_GUI_Menu()

editfield := Temp_File_Read(File_Install_Root_Folder,"TempAdd.txt")
Temp_File_Delete(File_Install_Root_Folder,"TempAdd.txt")

editfield2 := Temp_File_Read(File_Install_Root_Folder,"TempAdded.txt")
Temp_File_Delete(File_Install_Root_Folder,"TempAdded.txt")

TotalPrefixes := Temp_File_Read(File_Install_Root_Folder,"TempAmount.txt")
Temp_File_Delete(File_Install_Root_Folder,"TempAmount.txt")

If (editfield = "null") || (editfield2= "null") || (TotalPrefixes = "null")
{
	editfield =
	editfield2 =
	TotalPrefixes =
}
If(!Unit_test)
{
	Days := Calculate_Days_Since_Last_Update(updatestatus)
	If Days > %Updaterate%	; More than speficied days
	{
		Msg_Box_Result_Update := Move_Message_Box("4","Effectivity Macro Updater", "Would you like to check for a new update?" )
		lastupdate = %A_now%
		IniWrite, %updatestatus%,  %inifile%,update,updaterate
		IniWrite, %A_now%,  %inifile%, update,lastupdate
	}
	If Msg_Box_Result_Update = Yes
	Versioncheck()
}

Serials_GUI_Screen()

If Msg_Box_Result_Update = Yes
SetTimer, QuitBrowser, 1000

If (Unit_test)
	Checkvalues(Prefix_Store, First_Number_Set,  Second_Number_Set, "1")
return


;Sets the hotkey for Ctrl + 1 or Ctrl + numpad 1
$^Numpad1::
$^1::
{
	Formatted_Text := Format_Serial_Functions() ; Goes to the Formatserials subroutine
	Sort, Formatted_Text ; Sorts the prefixes in order
	Guicontrol,,Radio1,1
	Gui,Submit,NoHide
	gosub, radio_button
	Debug_Log_Event("Formatted text is " Formatted_Text)

	Formatted_Serial_Array := Object()

	Formatted_Serial_Array := Put_Formatted_Serials_into_Array(Formatted_Text)

	/* for testing********
	*/

	;~ SerialbreakquestionGUI() ; Goes to the Serialsgui.ahk and into the SerialbreakquestionGUI subroutine
	combine = 0
	Oneupserial = 1

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

		Guicontrol,1:, Editfield, %Editfield%*** ; Sets the listbox on teh GUi screen to the editfieldcombine vaariable and adds a newline
	}else  {
		Prefix_Count :=  Formatted_Text_Serial_Count(Formatted_Text)
		Guicontrol,1:, Editfield, %Formatted_Text%*** ; Sets the listbox on teh GUi screen to the editfieldcombine vaariable and adds a newline
	}

	totalprefixes = %Prefix_Count% ; Sets the totalprefixes variables to the Prefixcombinecount variable
	Guicontrol,, reloadprefixtext, There are a total of %totalprefixes% Serial Numbers to add to ACM ; Changes the valuse in the main GUI screen
	Winactivate, Effectivity Macro ; Make the Main GUi window  Active
	Guicontrol, Focus, Editfield ; Puts the cursor in the Editfield in teh Gui window
	send {Ctrl Down}{Home}{Ctrl Up} ; sends keystrokes to move the cursor to the top of the listbox
	Gui, Submit, NoHide ; Updates the Gui screen
	;~ gosub, ExportSerials
	return
}


Put_Formatted_Serials_into_Array(Formatted_Text)
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

Extract_Serial_Array(Combined_Serial_Array)
{
	editfield =
	For index, Element in  Combined_Serial_Array
	{
		;~ MsgBox % Combined_Serial_Array[A_Index]
		Editfield := Editfield element ",`n"
	}

	return Editfield
}

Formatted_Text_Serial_Count(Formatted_Text)
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

Copy_selected_Text()
{
	Clipboard =
	Send ^c ; sends a control C to the computer to copy selected text
	sleep(2)
	if clipboard =  ; if no text is seleted then clipboard will be remain blank
	return "No_Text_Selected"
	else
		return Clipboard ","
}

Format_Serial_Functions(Fullstring := "")
{
	Clear_Format_Variables()
	newline = `n
	sleep()
	If (Unit_test) ; For testing
	Fullstring := "621s (SN: 8KD1-00663,8KD00669,8KD00825,TRD00123, TRD00124-00165, TRD"
	Else
		FullString := Copy_selected_Text()

	If FullString = No_Text_Selected
	{
		Move_Message_Box("0","Error","Please ensure that text is selected before pressing Ctrl + 1.")
		Exit
	}


	Format_Removed_Text := Remove_Formatting(Fullstring) ; Goes to the Remove_Formatting funciton and stores the completed result into the Format_Removed_Text varialbe
		Debug_Log_Event("Format_Serials() Format_Removed_Text is " Format_Removed_Text)
MsgBox, % "Format removed is " Format_Removed_Text
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




Check_For_Single_Serials(PreFormatted_Text)
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



Combinecount(Prefix_Store_Array)
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

One_Up_All(Serial_Store_Array)
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


Combineserials(Formatted_Serial_Array)
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

Extract_Prefix(Serial_Number)
{
	StringMid, Prefix_Store, Serial_Number, 1, 3 ; Gets the prefix letters and stores them for later reuse
	return Prefix_Store
}

Extract_First_Set_Of_Serial_Number(Serial_Number)
{
	Stringmid, First_Half_Serial_Num,Serial_Number,4,5 ; takes the numbers after teh prefix and stores them into String%Counting%number variable
	return First_Half_Serial_Num
}

Extract_Serial_Dividing_Char(Serial_Number)
{
	Stringmid, Check_Char,Serial_Number,9,1 ; takes the next char after the first half of the serial numbers
	return Check_Char
}

Extract_Second_Set_Of_Serial_Number(Serial_Number)
{
	Stringmid, Second_Half_Serial_Num,Serial_Number,10,5
	return Second_Half_Serial_Num
}



Checkvalues(Prefix_Store, First_Number_Set,  Second_Number_Set, Reset := 0)
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




	Remove_Formatting(Selected_Text)
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


	PreFormat_Text(Format_Removed_Text)
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

	Prefix_Alone_Check_And_Add_One_UP(Text)
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


	add_digits(Serial_Number)
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


	Clear_Format_Variables()
	{
		global
		;Clear the variables
		Clipboard5 =
		FullString =
		Checkers =
		Clipboard =
		Clipboard1 =
		Checkeend  =
		Editfield =
		Parseclip =
		Guicontrol,, Editfield, %Editfield%
		formatfound =
		Stringthree =
		return
	}

/*
Below is the enter serials section
*/
Enter_Serials_Variable_Setup()
{
	global
   Prefixcount = 5
   addtime = 0
   Badlist =
    Stoptimer = 0
   Serialzcounter2 = 0
   Serialzcounter = 0
     StartTime := A_TickCount
 }

Enterallserials()
{
   global
   ; Initial setup before the macro starts to output to the screen

IniRead(inifile)

 Enter_Serials_Variable_Setup()

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






/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Below are the functions from the Autorun section Before it gets to Serials_GUI_Screen()i \./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/



	Folder_Exist_Check(Folder)
	{
		Result := FileExist("C:\" Folder)
		If Result =
		Result = Folder_Not_Exist
		else
			Result = Folder_Exist

		Debug_Log_Event("Folder_Exist_Check() .... " Folder "..... Result is " Result)

		return  Folder " - " Result
	}

	Folder_Create(Folder)
	{
		FileCreateDir, C:\%Folder%

		Debug_Log_Event("Folder_Create() .... " Folder "..... Result is " Result)
		sleep(5)
		return ErrorLevel
	}


	File_Exist_Check(File, Unit_test := 0)
	{
		If (Unit_test)
			Result := FileExist(A_Desktop "\" File)
		else
			Result := FileExist("C:\SerialMacro\" File)

		If Result =
		Result = File_Not_Exist
		else
			Result = File_Exist
		Debug_Log_Event("File_Exist_Check() ......C:\SerialMacro\" File "..... Result is " Result)
		return File " - "  Result
	}

	File_Create(File, Unit_test := 0)
	{
		;~ MsgBox %Unit_test% is unit test
		If (Unit_test)
			FileAppend,,%A_Desktop%\%File%
		else
			FileAppend,, C:\SerialMacro\%File%

		Debug_Log_Event("File_Create() ......C:\SerialMacro\" File)
		return ErrorLevel
	}



	Install_Requied_Files_Root( File_Install_Work_Folder)
	{
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\How to use Effectivity Macro.pdf, %File_Install_Work_Folder%\How to use Effectivity Macro.pdf,1
		return errorlevel
	}

	Install_Requied_Files_Icons( File_Install_Work_Folder)
	{
		Problems = 0
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\icons\serial.ico, %File_Install_Work_Folder%\icons\serial.ico,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\icons\paused.ico, %File_Install_Work_Folder%\icons\paused.ico,1
		If (Errorlevel)
			Problems = 1

		return Problems
	}

	Install_Requied_Files_Images( File_Install_Work_Folder)
	{
		Problems = 0
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\red_image.png, %File_Install_Work_Folder%\images\red_image.png,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\active_plus.png, %File_Install_Work_Folder%\images\active_plus.png,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\orange_button.png, %File_Install_Work_Folder%\images\orange_button.png,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\paused.png, %File_Install_Work_Folder%\images\paused.png,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\start.png, %File_Install_Work_Folder%\images\start.png,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Running.png, %File_Install_Work_Folder%\images\Running.png,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\Stopped.png, %File_Install_Work_Folder%\images\Stopped.png,1
		If (Errorlevel)
			Problems = 1
		FileInstall, C:\Users\karnijs\Desktop\Autohotkey\02_Effectivity Macro\1.4\Install_Files\images\background.png, %File_Install_Work_Folder%\images\background.png,1
		If (Errorlevel)
			Problems = 1

		return Problems
	}

	/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Below is the Create_Tray_Menu funciton along with the functions it solely calls to .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

	Create_Tray_Menu()
	{
		Menu, Tray, NoStandard
		Menu, Tray, Add, How to use, HowTo
		Menu Tray, Add, Check For update, Versioncheck
		Menu, Tray, Add, Quit, Quitapp
		return
	}


	Howto()
	{
		splashtexton,,Effectivity Macro, Loading PDF
		Run, C:\SerialMacro\How to use Effectivity Macro.pdf
		sleep(20)
		SplashTextOff
		return
	}

	Quitapp:
	{
		Result := 	Move_Message_Box("262148","Quit " Effectivity_Macro, "Are you sure you want to quit?")

		If result =  Yes
		{
			Stopactcheck = 1
			Gui 1: -AlwaysOnTop
			Gui_Image_Show("Stopped") ; Options are Start, Paused, Running, Stopped
			Gui, Submit, NoHide
			Send {Shift Up}{Ctrl Up}
			breakloop = 1
			ExitApp
		}
		Return
	}

	/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Below is the Create_Main__gui_Menu funciton along with the functions it solely calls to .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

	Create_Main_GUI_Menu()
	{
		Menu, BBBB, Add, &Check For Update , Versioncheck
		Menu, BBBB, Add, &Options, OptionsGui
		Menu, BBBB, Add,
		Menu, CCCC, Add, &Run							(Crtl + 2), Enterallserials
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
			Gui_Image_Show("Paused") ; Options are Start, Paused, Running, Stopped
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
			Gui_Image_Show("Running") ; Options are Start, Paused, Running, Stopped
			Gui, Submit, NoHide
			Pause, toggle, 1
		}
		return
	}

	Exit_Program(Unit_Test := 0)
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

	restart_macro()
	{
		Result := Move_Message_Box("262148",Effectivity_Macro, "Are you sure that you want to reload the program?" )

		If Result =  yes
		Reload

		return
	}

	restart_macro_Effectivity()
	{
		global nextserialtoadd, EditField, EditField2, TempSavefile,totalprefixes,Serialcount

		Result := Move_Message_Box("262148",Effectivity_Macro, "Are you sure that you want to reload the program?" )

		If Result =  yes
		{
			GuiControlGet, nextserialtoadd
			GuiControlGet, EditField
			GuiControlGet, EditField2
			If nextserialtoadd =
			TempSavefile = %editfield%
			else
				TempSavefile = %nextserialtoadd%`,`n%editfield%

			FileAppend, %TempSavefile%, C:\SerialMacro\TempAdd.txt
			FileAppend, %EditField2%, C:\SerialMacro\TempAdded.txt
			FileAppend, %totalprefixes%, C:\SerialMacro\Tempamount.txt
			FileAppend, %Serialcount%, C:\SerialMacro\Tempcount.txt
			Reload
		}

		return
	}


	/*
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\. Gui Screens .\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\..\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.\./.\./.\.
	*/

	Serials_GUI_Screen()
	{
		Global

		activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

		If totalprefixes < 1
		{
			TotalPrefixes = 0
		}
		gui 1:add, Edit, x10 y50 w390 h240  vEditField,%editfield%
		gui 1:add, Edit, xp yp w390 h240 vEditField2,%editfield2%

		Gui 1:Add, Picture, x315 y310 w50 h50 +0x4000000  BackGroundTrans vStarting gstartmacro , C:\SerialMacro\images\Start.png
		Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vRunning, C:\SerialMacro\images\Running.png
		Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vpaused  gpausesub, C:\SerialMacro\images\Paused.png
		Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vStopped grestart_macro, C:\SerialMacro\images\Stopped.png
		Gui, 1:Add, Picture, x0 y0 w410 h400 +0x4000000 , C:\SerialMacro\images\background.png

		Gui 1:Add, Edit, xp+165 yp+343 w110 h20  vnextserialtoadd, %nextserialtoaddv%

		Gui 1:Add, Text, x5 y5 w300 h25 BackgroundTrans +Center vreloadprefixtext, There are a total of %TotalPrefixes% Effectivity to add to ACM

		Gui 1:add, Radio, xp+25 yp+25 w130 h20 BackGroundTrans Checked vradio1 gradio_button, Effectivity to be added
		Gui 1:add, Radio, xp+155 yp w140 h20 BackGroundTrans  vradio2 gradio_button, Effectivity already added

		Gui 1:Add, Text, xp-170 Yp+265 W250 h13 BackGroundTrans vserialsentered, Number of Effectivity successfully added to ACM = %Serialcount%

		Gui 1:Add, Text, Xp Yp+15 w250 h13  BackGroundTrans , If macro is operating incorrectly, press Esc to reload
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

	Editfield_Control(Textbox, Gui_Number := 1)
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

	startmacro:
	{
		skipbox = 0
		enterallserials()
		return
	}

	Enterallserials()
	{
		Msgbox, enterallserials
		return
	}

	Gui_Image_Show(Image)
	{
		If ( image = Paused)
		{
			Guicontrol,hide, Start
			Guicontrol,show, paused
			Guicontrol,hide, Stopped
			Guicontrol,hide, Running
		}

		If (image = Running)
		{
			Guicontrol,Show, Start
			Guicontrol,Hide, paused
			Guicontrol,hide, Stopped
			Guicontrol,hide, Running
		}
		If (image = Stopped)
		{
			Guicontrol,hide, Start
			Guicontrol,hide, paused
			Guicontrol,Show, Stopped
			Guicontrol,hide, Running
		}
		If (image = start)
		{
			Guicontrol,show, Start
			Guicontrol,hide, paused
			Guicontrol,hide, Stopped
			Guicontrol,hide, Running
		}
		return
	}


	Temp_File_Read(File_Install_Root_Folder,File_Name)
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

	Temp_File_Delete(File_Install_Root_Folder,File_Name)
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

			SerialbreakquestionGUI()
			{
				global createexcel
				activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

				gui 1: -alwaysontop
				Gui, 8:Add, Picture, x0 y0 w400 h90 +0x4000000, %File_Install_Work_Folder%\images\background.png
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


			oneup:
			{
				UnPausescript()
				GuiControlGet, checked,, createexcel
				gui 1: +alwaysontop
				Oneupserial = 1
				Gui, 8:destroy
				Return
			}

			combinequstion:
			{
				UnPausescript()
				GuiControlGet, checked,, createexcel
				gui 1: +alwaysontop
				Gui, 8:destroy
				combine = 1
				Return
			}

			keepseperated:
			{
				UnPausescript()
				GuiControlGet, checked,, createexcel
				gui 1: +alwaysontop
				Gui, 8:destroy
				combine = 0
				Oneupserial = 0

				Return
			}

			aboutmacro()
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

			emaillink()
			{
				Run,  mailto:Karnia_Jarett_S@cat?Subject=Effectivity Macro
				return
			}

			boxlink()
			{
				Run, %Program_Location_Link%
				return
			}


			sleep(Amount := 1)
			{
				amount := amount * 100
				Sleep %Amount%
				Return
			}


			Calculate_Days_Since_Last_Update(updatestatus)
			{
				Today := A_Now		; Set to the current date first
				EnvSub, Today, %updatestatus%, Days 	; this does a date calc, in days
				Return Today
			}

			Versioncheck()
			{
				global Checkversion, Update_Check_URL, Program_Location_Link, Version_Number
				Load_ini_file(inifile)
				Progress,  w200, Updating..., Gathering Information, Effectivity Macro Updater
				Progress, 0
				sleep(2)
				Versioncount = 0
				settimer, versiontimeout, 500
				create_checkgui(hwnd, ParentGUI, wb)
				Progress,  w200, Updating..., Fetching Server Information, Effectivity Macro Updater
				Progress, 15

				DllCall("SetParent", "uint",  hwnd, "uint", ParentGUI)
				wb.Visible := True
				WinSet, Style, -0xC00000, ahk_id %hwnd%

				Progress, 25
				sleep(2)
				;~ MsgBox, %Update_Check_URL%
				Update_Check_URL := "https://docs.google.com/document/d/1woiaqcTjqkABrIecRERDAt6nqiEknFWdySqRmie7bCM/edit?usp=sharing"
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
				;~ Progress, Off


				Doc_Title := Check_Doc_Title()

				update_Version:=  Format_Serial_Check_Title(Doc_Title)


				;~ msgbox, Checkversion is %Checkversion%
				If update_Version = Not_Found
				{
					Progress,  w200,Updating..., Error Occured. Update Not Able To Complete, Effectivity Macro Updater
					Progress, 0
					sleep()
					Progress, off
					Move_Message_Box("0","Effectivity Macro Updater", "Error `n Cannot find Server" )
					return
				}

				If update_Version <= %Version_Number%
				{
					Progress,  w200,Updating..., Macro is Up to date., Effectivity Macro Updater
					Progress, 100
					sleep(10)
					Progress, off
					settimer, versiontimeout, Off
					;Msgbox,,Serial Macro Updater,Macro is Up to date.
				}

				If update_Version > %Version_Number%
				{
					settimer, versiontimeout, Off

					Result := Move_Message_Box("262148","Effectivity Macro Updater", " New update found. Would you like to open the Cat Box site to download the latest version?" )
					;~ MsgBox, %Program_Location_Link%
					If Result =  yes
					Run, %Program_Location_Link%
				}

				IniWrite, %updaterate%, %inifile%,update,updaterate
				IniWrite, %A_now%,  %inifile%, update,lastupdate
				Progress, w200,,Disconnecting From Server..., Effectivity Macro Updater
				Progress,25
				Sleep()
				Progress, 50

				Gui,2:Destroy
				wb.Quit
				wb:= ""
				Progress, Off
				return
			}

			QuitBrowser:
			{
				SetTimer, QuitBrowser, Off

				return
			}

			Check_Doc_Title()
			{
				Result = Not_Found
				Loop, 3
				{
					Winactivate, Serial version
					wingettitle, Google_Doc_Title, A
					sleep(10)
					;~ Msgbox, Title is  %Google_Doc_Title%
					Result = Not_Found
					If Google_Doc_Title !=
					{
						Result := Google_Doc_Title
						break
					}}
					return Result
				}

				Format_Serial_Check_Title(Title)
				{
					If Title != Not_Found
					{
						StringGetPos, pos, Title, #, 1
						Update_Check_Version_Number := SubStr(Title, pos+2)
						Update_Check_Version_Number := SubStr(Update_Check_Version_Number,1,3)
						return Update_Check_Version_Number
					}
					else
						return "Not_Found"
				}


				create_checkgui(ByRef hwnd, ByRef ParentGUI, ByRef wb)
				{
					global
					wb := ComObjCreate("InternetExplorer.Application")
					sleep(5)
					Wb.AddressBar := false
					;~ Wb.AddressBar := true
					wb.MenuBar := false
					;~ wb.MenuBar := true
					wb.ToolBar := false
					;~ wb.ToolBar := true
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
					;~ Gui,2:show, x50 y50 W1000 h1000 , Updater
					return
				}

				versiontimeout:
				{
					Versioncount++
					If Versioncount = 60
					{
						Progress,  w200,Updating..., Error Occured.Cannot connect to server for update check, please check for internet connection., Effectivity Macro Updater
						Progress, 0
						sleep(2)
						Progress, Off
						SplashTextOff
						Gui,2:Destroy
					}
					Return
				}

				OptionsGui:
				{
					activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

					Load_ini_file(inifile)
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
					Write_ini_file(inifile)
					gui 10:destroy
					return
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

		Debug_Log_Event(Event)
		{
			global Log_Events

			If (Log_Events)
			{
			OutputDebug, %Event%
			Sleep(.5)

	}
	return
}
#`::ListLines
+#~::ListVars