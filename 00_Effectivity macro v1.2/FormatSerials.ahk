
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

Guicontrol,, reloadprefixtext, There are a total of %totalprefixes% Serial Numbers to add to ACM
;sleep 200
Winactivate, Enter Serial Macro
Guicontrol, Focus, Editfield
;msgbox, pause
Send {Ctrl Down}{End}{Ctrl Up}
Sleep 300

Send {BackSpace}{*}{*}{*}
send {Ctrl Down}{Home}{Ctrl Up}
gosub, Radio1h


Gui, Submit, NoHide
Return
}

/*
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\ combineserials Subroutine /\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
*/

combineserials:
{
counting = 0
Match = 0

;msgbox, clipboard6 is `n`n%clipboard6%

Loop, Parse, clipboard6, `n,all
{
;msgbox, combineserials loop is `n%A_LoopField%
counting++
StringMid, String%Counting%prefix, A_loopfield, 1, 3
Checkprefix := String%Counting%prefix
If checkprefix = `,
{
;msgbox, nothing there
checkprefix = 
lastnums = 
Continue
}

gosub, matchprefix

Stringmid, String%Counting%number,A_LoopField,4,5
Stringmid, String%Counting%check,A_LoopField,9,1
endchar :=  String%Counting%check
begnumcheck := String%Counting%number
;msgbox, endchar is `n%endchar%
	If endchar = `-
	{
	Stringmid, String%Counting%last,A_LoopField,10,5
	Lastnums :=  String%Counting%last
	}
	If Endchar = `,
	{
	Endchar = `-
	lastnums = %begnumcheck%
	}
	

;msgbox, Prefix is `n%Checkprefix% `n`n mischeck is `n%begnumcheck% `n`n checkchar is `n%endchar% `n`n lastnums is `n%lastnums%

If Match = 1
{
Gosub, Checkvalues
Prefix%checkprefix% = %Checkprefix%%begnumcheck%%endchar%%lastnums%	
Serialnumzz := Prefix%Checkprefix%
;Clipboard7 = %clipboard7%%Serialnumzz%
Match = 0
Continue
}

Else if Match = 0
{
Prefix%checkprefix% = %Checkprefix%%begnumcheck%%endchar%%lastnums%
}
;Serialnumzz = %Checkprefix%%begnumcheck%%endchar%%lastnums%`, 
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
	}
}

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
	}
	Else
{
	Clipboard1 =  %clipboard1%%NewStr5%
	}
}
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
}
}
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
