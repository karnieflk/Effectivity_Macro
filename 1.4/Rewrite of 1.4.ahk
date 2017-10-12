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
Global Prefix_Number_Location_Check, First_Effectivity_Numbers, Title, sleepstill, Current_Monitor, Log_Events, unit_test, File_Install_Work_Folder, Oneupserial, combineser

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

inifile = c:\Serialmacro\config.ini

File_Install_Root_Folder = C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files

File_Install_Work_Folder = C:\SerialMacro
      
Image_Red_Exclamation_Point = %File_Install_Work_Folder%\red_image.png
IMage_Actve_Add_Button = %File_Install_Work_Folder%\Active_plus.png
Image_Active_Apply_Button = %File_Install_Work_Folder%\orange_button.png

Unit_Test = 0 ; Set this to 1 to perform unit tests and logging. 
Log_Events = 0 ;Set this to 1 to perform logging

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
  

  
Install_Requied_Files_Root(File_Install_Work_Folder)
Install_Requied_Files_Icons(File_Install_Work_Folder)
Install_Requied_Files_Images(File_Install_Work_Folder)

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

Msg_Box_Result := Update_Check(updatestatus)
If Msg_Box_Result = Yes
	Versioncheck()

Serials_GUI_Screen()
   
return

  ;Sets the hotkey for Ctrl + 1 or Ctrl + numpad 1
$^Numpad1::
$^1::
{
  Formatted_Text := Format_Serials() ; Goes to the Formatserials subroutine
    Sort, Formatted_Text ; Sorts the prefixes in order
    Guicontrol,,Radio1,1
    Gui,Submit,NoHide
    gosub, radio_button
SerialbreakquestionGUI() ; Goes to the Serialsgui.ahk and into the SerialbreakquestionGUI subroutine
If Oneupserial = 1
{
Prefixmatching := Combineserials(Formatted_Text) ;goes to the combine Serials subroutine
CombineCount(Prefixmatching)
}
If combine = 1
Combineserials(Formatted_Text) ;goes to the combine Serials subroutine  
 
     
   return
}

Copy_selected_Text()
{
   Send ^c ; sends a control C to the computer to copy selected text
   sleep(2)
   if clipboard =  ; if no text is seleted then clipboard will be remain blank
   return "No_Text_Selected"
   else
   return Clipboard ","
}

Format_Serials()
{
   global
   Clear_Format_Variables()
   newline = `n
   sleep()
   FullString := Copy_selected_Text()
If FullString = No_Text_Selected
   {
      Move_Message_Box("0","Error","Please ensure that text is selected before pressing Ctrl + 1.")
      Exit
   }

     ;~ Checkers = %Fullstring% ; Sets the checkers vairable to the contents of the fullstring variable
   Format_Removed_Text := Remove_Formatting(Fullstring) ; Goes to the Remove_Formatting funciton and stores the completed result into the Format_Removed_Text varialbe
   
   PreFormatted_Text = PreFormat_Text(Format_Removed_Text) ; Goes to the PreFormat_Text function and stors the completed result into the Preformatted_text variable
   ;msgbox, CLipboard1 is `n`n %clipboard1%
   
   Serial_Add_Count := Prefix_Number_Checking(PreFormatted_Text) ;  Take info from the PreFormatted_Text variable and checks to see of it it just the serial with no numbers attached to it. If is, then adds 00001-99999 to prefix. Also changes the Parseclip variable to the Number of non combined Serials   
   
   ;msgbox, CLipboard5 is `n`n %clipboard5%
   Formatted_Text = %PreFormatted_Text%%newline% ; makes the editfieldbreaks variable the same as the PreFormatted_Text variable contents and adds a return carriage to it.
   return Formatted_Text
}
   
   
Still refactor below - used to be part of format serials After the combine serials subroutine
   
   ;msgbox, prefixlist is `n`n%Prefixlist% ; FOr diagnostics
   Combinecount(Prefixmatching)
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
    Sort, EditfieldCombine
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
   
   Send {*}{*}{*} ; adds 3 astriks
   send {Ctrl Down}{Home}{Ctrl Up}{del} ; sends keystrokes to move the cursor to the top of the listbox
   gosub, Radio1h ; Goes to the Radio1h subroutine
   
   
   Gui, Submit, NoHide ; Updates the Gui screen
   gosub, ExportSerials
   Return
}

Combineserials(Formatted_Text)
{
   global 
   
   Loop, Parse, Formatted_Text, `n,all ; loop to divide the Formatted_Text variable by the carraige returns
   {
      ;msgbox, combineserials loop is `n%A_LoopField%
  
      Prefix_Store := Extract_Prefix(A_LoopField)
      
      If Prefix_Store = `, ; checks to see if the Prefix_Store variable is a comma
      {
         ;msgbox, nothing there
         Prefix_Store = ; sets the Prefix_Store variable to nothing
         Sencond_Number_Set =  ; sets the Sencond_Number_Set variable to nothirng
         Continue ; skips over the rest of the loop and starts at the top of the parse loop
      }
      
     Match_result := matchprefix(Prefix_Store, Prefixmatching) ;goes to the matchprefix function
      
      First_Number_Set := Extract_First_Set_Of_Serial_Number(A_LoopField)
   Middle_Char := Extract_Serial_Dividing_Char(A_LoopField)
          
      If Middle_Char = `- ; checks if Middle_Char variable is a hyphen
             Sencond_Number_Set := Extract_Second_Set_Of_Serial_Number(A_LoopField)
            
   else If Middle_Char = `, ; if the Middle_Char variable is a comma
      {
         Middle_Char = `- ; makes the Middle_Char variable a hyphen
         Sencond_Number_Set = %First_Number_Set% 
      }

If Match_Result = Already_Matched 
         Checkvalues(Prefix_Store,First_Number_Set,Sencond_Number_Set) ; goes to the Checkvalues subroutine	     

   Prefix%Prefix_Store% = %Prefix_Store%%begnumcheck%%endchar%%lastnums% ;makes the Prefix%Prefix_Store% variable be the combination of the all those other variables
         }
         
   Return Prefixmatching
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

Checkvalues(Prefix_Store,ByRef First_Number_Set, ByRef Sencond_Number_Set)
{
   
   oldprefix := Prefix%Prefix_Store%
   ;msgbox, oldprefix is `n %oldprefix%
 Prefix_Beg :=  Extract_First_Set_Of_Serial_Number(oldprefix)
 Prefix_Last :=  Extract_Second_Set_Of_Serial_Number(oldprefix)

   
   If 	Prefix_Beg > %Sencond_Number_Set%
   {
      Sencond_Number_Set = %Prefix_Beg%
   }
   
   If 	Prefix_Beg < %First_Number_Set%
   {
      First_Number_Set = %Prefix_Beg%
   }
   
   
   If 	Prefix_Last > %Sencond_Number_Set%
   {
      Sencond_Number_Set = %Prefix_Last%
   }	
      
   return
}



matchprefix(Prefix_Store, ByRef Prefixmatching)
{
   ;StringReplace, Prefixmatching,prefixmatching,%A_Space%,,all
   Loop, Parse, Prefixmatching,`,
   {
      ;msgbox, loopfield match is %A_LoopField%
      If a_loopfield = 
           Continue     
      
   else IF A_LoopField = %Prefix_Store%
		   Return "Already_Matched"
   }
   Prefixmatching = %Prefixmatching%%Prefix_Store%`,
   return "Updated Prefixmatching with " Prefix_Store
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
   
   Return Selected_Text
}


PreFormat_Text(Format_Removed_Text)
{
   Counter = 1
   
   ;Loops Format_Removed_Text variable to clean up the all the entereed text. Removes carriage returns with parse, removes any spaces, changes the ) to double , 
   ; Changes 1-up to nothing, changes -up to 999999, changes ), to nothing
   Loop, Parse, Format_Removed_Text, `r`n
   {
      StringGetPos, pos, A_loopfield, :, 1	
      String%Counter% := SubStr(A_LoopField, pos+2)
      Format_Text := String%Counter%

      StringReplace, Format_Text,Format_Text,%A_Space%,,All	  
      StringReplace, Format_Text,Format_Text, `),`,,All    
      StringReplace, Format_Text,Format_Text, 1`-up,,All  
      StringReplace, Format_Text,Format_Text, `-up,`-99999, All
      StringReplace, Format_Text,Format_Text, `,,`,`n, All	

      Counter++
      If Format_Text = 
      {
         continue
      }else  {
         Full_Text =  %Full_Text%%Format_Text%
      }}
   
   Full_Text = %Full_Text%`,
   ;Msgbox, Before replace comma %clipboard1%
   StringReplace, Full_Text,Full_Text, `,,,All

   Return Full_Text
}

Prefix_Number_Checking(Byref Text)
{
   Serial_Add_Count = 0	
   Counter = 0
   ;Msgbox, parseclip is `n`n %parseclip%
   ;msgbox parseclip is `n%parseclip%
   Loop, parse, Text, `n
   {
      ;msgbox, parseclip loop is `n`n %A_LoopField%
      If a_loopfield = 
      {
         Continue
      }
      Serial_Add_Count++	
      Counter++
      add_comma = %A_LoopField%`,
      StringGetPos, pos, add_comma, `,, 1
      If pos = 3
      {
         StringReplace, Attach_Serial_Numbers,add_comma,`,,00001`-99999`,`n,all
         ;msgbox, if 3 add 1-up is:`n%Clippyt%
      }
      Else if pos != 3
      {
         StringMid, String%Counter%_prefix, A_loopfield, 1, 3
         Prefixstore := String%Counter%_prefix
         ;Prefixlist = %Prefixlist%`,%Prefixstore%
         
         ;msgbox, prefix is %Prefixstore%
         StringTrimLeft, String%Counter%left, A_loopfield, 3
         StringTemp_Left := String%Counter%left
         ;msgbox, StringTemppre  is %StringTemppre%
         StringTemp_end := add_digits(StringTemp_Left)
         ;msgbox, StringTemp  is %StringTempend%
         Stringtemp = %Prefixstore%%StringTemp_end%
         StringReplace, StringTemp,StringTemp,`),`,`n,all
         ;msgbox, %stringTemp%
         Attach_Serial_Numbers = %Stringtemp%

      }
      Text = %Text%%Attach_Serial_Numbers%
      }
 
   return Serial_Add_Count
}

add_digits(Serial_Number)
{
   Final_Combined_Digits =
   combinestring = 
   Loop, Parse, Serial_Number, `-
   {
      StringReplace, Remove_End_Parenthesis,A_LoopField,`),,
      int = %Remove_End_Parenthesis%
      Loop, % 5-StrLen(int)
      int = 0%int%
      combinestring = %combinestring%`-%Int%
   }
   combinestring := SubStr(combinestring, 2)
   Final_Combined_Digits = %combinestring%`)
   ;msgbox, end of add zeros is  %Stringtempend%
   
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
        If Log_Events = 1
         FileAppend, Folder_Exist_Check....%Folder%.....Result = %Result% `n, %A_Desktop%\Serial_Macro_Log_File.txt
	return  Folder " - " Result
	}

Folder_Create(Folder)
	{
		 FileCreateDir, C:\%Folder%
         If Log_Events = 1
         FileAppend, Folder_Create....%Folder%.....Result = %Result% `n, %A_Desktop%\Serial_Macro_Log_File.txt
	   sleep(5)	
	return       
	}


File_Exist_Check(File)
	{
		Result := FileExist("C:\SerialMacro\" File)
        If Result = 
         Result = File_Not_Exist
         else
            Result = File_Exist
        If Log_Events = 1
         FileAppend, File_Exist_Check....C:\SerialMacro\%File%.....Result = %Result% `n, %A_Desktop%\Serial_Macro_Log_File.txt
         
	return File " - "  Result
	}

	File_Create(File)
	{
      		FileAppend, "C:\SerialMacro\" %File%		
              If Log_Events = 1
         FileAppend, File_Create....C:\SerialMacro\%File% `n, %A_Desktop%\Serial_Macro_Log_File.txt
	return 
	}



Install_Requied_Files_Root( File_Install_Work_Folder)
	{
   	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\How to use Effectivity Macro.pdf, %File_Install_Work_Folder%\How to use Effectivity Macro.pdf,1
	return
	}
    
    Install_Requied_Files_Icons( File_Install_Work_Folder)
	{
   
   	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\icons\serial.ico, %File_Install_Work_Folder%\icons\serial.ico,1
   	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\icons\paused.ico, %File_Install_Work_Folder%\icons\paused.ico,1
     
      return
     }
     
         Install_Requied_Files_Images( File_Install_Work_Folder)
	{
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\red_image.png, %File_Install_Work_Folder%\images\red_image.png,1
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\active_plus.png, %File_Install_Work_Folder%\images\active_plus.png,1
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\orange_button.png, %File_Install_Work_Folder%\images\orange_button.png,1
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\paused.png, %File_Install_Work_Folder%\images\paused.png,1
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\start.png, %File_Install_Work_Folder%\images\start.png,1
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\Running.png, %File_Install_Work_Folder%\images\Running.png,1
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\Stopped.png, %File_Install_Work_Folder%\images\Stopped.png,1
      	FileInstall, C:\Users\karnijs\Desktop\Autohotkey\Effectivity Macro\1.4\Install_Files\images\background.png, %File_Install_Work_Folder%\images\background.png,1
      
      return
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
    
   Exit_Program()
{
   global Serialcount
 Result := Move_Message_Box("262148",Effectivity_Macro, " The number of successful Serial additions to ACM is %Serialcount% `n`n Are you sure you want to Stop the macro?.`n`n Press YES to stop the Macro.`n`n No to keep going.")	   

If Result = Yes
   {
      Stopactcheck = 1						
      Gui 1: -AlwaysOnTop
      Gui_Image_Show("Stopped")
      Send {Shift Up}{Ctrl Up}
      breakloop = 1
      Exit
}
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
     If Log_Events = 1
      FileAppend, Fileread....%File_Name%.....Variable_Store %Variable_Store% `n , %A_Desktop%\Serial_Macro_Log_File.txt
	FileDelete, %File_Install_Root_Folder%\%File_Name%  
    If Log_Events = 1
      FileAppend, FileDelete....%File_Name% `n, %A_Desktop%\Serial_Macro_Log_File.txt
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
    If Log_Events = 1
      FileAppend, FileDelete....%File_Name% `n, %A_Desktop%\Serial_Macro_Log_File.txt	IfExist, %File_Install_Root_Folder%\%File_Name%
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
IfMsgBox = Ok
   Result = Ok
IfMsgBox = Cancel
   Result = Cancel
IfMsgBox = Abort
   Result = Abort
IfMsgBox = Ignore
   Result = Ignore
IfMsgBox = Retry
   Result = Retry
IfMsgBox = continue
   Result = continue
IfMsgBox = TryAgain
   Result = TryAgain
IfMsgBox= Timeout
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
   combineser = 1
     Return
}

keepseperated:
{
   UnPausescript()
   GuiControlGet, checked,, createexcel
   gui 1: +alwaysontop
   Gui, 8:destroy
   combineser = 0
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

Update_Check(updatestatus)
{
	global
    Load_ini_file(inifile)
    Today := A_Now		; Set to the current date first
      EnvSub, Today, %updatestatus% , Days 	; this does a date calc, in days
      If Today > %Updaterate%	; More than speficied days
      {
      Result := Move_Message_Box("4","Effectivity Macro Updater", "Would you like to check for a new update?" )    
      lastupdate = %A_now%
      Write_ini_file(inifile)
			;~ IniWrite, %updatestatus%,  %inifile%,update,updaterate	
            ;~ IniWrite, %A_now%,  %inifile%, update,lastupdate
      return Result 
  }}
  
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
  
  
   Checkversion := Check_Doc_Title()
   wb.Close
    ;~ msgbox, Checkversion is %Checkversion%
   If Checkversion = Not_Found
		{
		Progress,  w200,Updating..., Error Occured. Update Not Able To Complete, Effectivity Macro Updater
		Progress, 0
		 sleep()
		 Progress, off		
		 Move_Message_Box("0","Effectivity Macro Updater", "Error `n Cannot find Server" )
		 return
		  }
   
   
   If Checkversion <= %Version_Number%
   {
      Progress,  w200,Updating..., Macro is Up to date., Effectivity Macro Updater
      Progress, 100
      sleep()
      Progress, off
      settimer, versiontimeout, Off
      ;Msgbox,,Serial Macro Updater,Macro is Up to date.
   }
   
   If Checkversion > %Version_Number%
   {
	  settimer, versiontimeout, Off
      
      Result := Move_Message_Box("262148","Effectivity Macro Updater", " New update found. Would you like to open the Cat Box site to download the latest version?" )
      ;~ MsgBox, %Program_Location_Link%
      If Result =  yes 
            Run, %Program_Location_Link%
   }
      
	IniWrite, %updaterate%, %inifile%,update,updaterate	
	IniWrite, %A_now%,  %inifile%, update,lastupdate
    Progress, Off
   Gui,2:Destroy
   return
}

Check_Doc_Title()
{
   Result = Not_Found
   Loop, 3
{
	Winactivate, Serial version
   wingettitle, Google_Doc_Title, A
   sleep()
   ;~ Msgbox, Title is  %title% 
   
   If Google_Doc_Title != 
   {
      Result = 
  	 break
 }}
  Progress, Off
If Result != Not_Found
{
StringGetPos, pos, Google_Doc_Title, #, 1	
   Update_Check_Version_Number := SubStr(Google_Doc_Title, pos+2)
   Result := SubStr(Update_Check_Version_Number,1,3)
 }

     
 return Result
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
    Result := Move_Message_Box("262148","Quit" Effectivity_Macro, "Are you sure you want to quit?" )
      ;~ MsgBox, %Program_Location_Link%
      If Result =  yes 
         Quitapp
      
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

#`::ListLines
+#~::ListVars
