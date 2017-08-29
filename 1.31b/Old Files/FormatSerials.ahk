;Sets the hotkey for Ctrl + 1 or Ctrl + numpad 1
$^Numpad1::
$^1::
{
   Gosub, FormatSerialsMacro ; Goes to the Formatserials subroutine
   return
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
   sleep 100
   Send ^c ; sends a control C to the computer to copy selected text
   sleep 200
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
   ;sleep 200
   Winactivate, Effectivity Macro ; Make the Main GUi window  Active
   Guicontrol, Focus, Editfield ; Puts the cursor in the Editfield in teh Gui window
   ;msgbox, pause
   Send {Ctrl Down}{End}{Ctrl Up} ; Send the keystrokes so the cursor is at the bottom of the window
   Sleep 300 ; delays the script for 300 miliseconds
   
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
sleep 100
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
sleep 500
Progress, off
;MsgBox, 64,, Done!, 1
;run, %WorkbookPath%
Return
}

CombineStraglers()
{


return
}