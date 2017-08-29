#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
inifile = c:\arbortextmacros\config.ini 

Load_ini_file(inifile)
Sleep 100
GUI_ini_file(inifile)
Return








Load_ini_file(inifile)
 {
	global
	
	Ini_var_store_array:= Object()
	Section_store_array:= Object()
		loop,read,%inifile%
  {
	If A_LoopReadLine = 
		continue
	
    if regexmatch(A_Loopreadline,"\[(.*)?]")
      {
        Section :=regexreplace(A_loopreadline,"(\[)(.*)?(])","$2")
		StringReplace, Section,Section, %a_space%,,All
		Section_store_array.Insert(Section)
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
      }
  }
	
	return
}

GUI_ini_file(inifile)
{
	global
	Yvar=30
	tabcount = 1
	
	for index, element in Section_store_array 
	{
		If index = 1
			Tabs :=  element
		else
		Tabs := Tabs "|"  element
}
;~ MsgBox, %Tabs%
		Gui, Add, Tab2,  w749 h450, %tabs%
			GUi, Add, button, y470 x10 gsave_settings, Save config File
	for index, element in Ini_var_store_array 
	{	
		
	
	StringSplit, INI_Write,element, `:
	Varname := INI_Write1
		
		If INI_Write2 = %elementTemp%
		{
			Gui, add, Text, x10 yp+%Yvar% w210 ,%INI_Write1% is set to 
	GUi, add, Edit, x215 yp v%INI_Write1%, % %INI_Write1%
	}
			else 
		{
			;~ MsgBox, elementTemp is %elementTemp%`n INI_Write2 is %INI_Write2%
			elementTemp = %INI_Write2%
			GUi, tab, %ini_write2%
				Gui, add, Text, x10 y30 w210 ,%INI_Write1% is set to 
			GUi, add, Edit, x215 yp v%INI_Write1%, % %INI_Write1%
			}

	    } 
	

	Gui, show, w750 h500, 
	return
}

Save_Settings()
{
	global
	gui,Submit,NoHide
	Write_ini_file(inifile)
	MsgBox, saved
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