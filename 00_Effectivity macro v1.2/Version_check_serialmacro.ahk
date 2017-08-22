Updatechecker:
{
	if not (FileExist("C:\Serialmacro\Settings.ini"))
	{
		FileAppend,, C:\Serialmacro\Settings.ini
		activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Do you want to check for an update"
Settimer, winmovemsgbox, 20

		Msgbox,4,%A_Space%Effectivity Macro Updater, %A_Space%Looks like This may be the first time running this application. Do you want to check for an update?
		ifmsgbox Yes
		{
		gosub, Versioncheck		
		}
		Else
		{
		IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
		IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
		Return
		}
	}

	if (FileExist("C:\Serialmacro\Settings.ini"))
	{
	IniRead, updatestatus, C:\Serialmacro\Settings.ini, update,lastupdate
	IniRead, Updaterate, C:\Serialmacro\Settings.ini, update,updaterate
	NumberOfDays := A_Now		; Set to the current date first
	EnvSub, NumberOfDays, %updatestatus% , Days 	; this does a date calc, in days
	If NumberOfDays > %Updaterate%	; More than 13 days
		{
				activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := "Would you like to check for a new update"
Settimer, winmovemsgbox, 20
		MsgBox,4,%A_Space%Effectivity Macro Updater, It has been %NumberOfDays% days since the last update check.`n`n Would you like to check for a new update?`n`n
		ifmsgbox Yes	
			gosub, Versioncheck
			else
			{
			IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
			IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
			}
		}
	}
Return
}
	
	
Versioncheck:
{
Progress,  w200, Updating..., Gathering Information, Effectivity Macro Updater
Progress, 0
sleep 200
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
sleep 200

wb.navigate("https://docs.google.com/document/d/1woiaqcTjqkABrIecRERDAt6nqiEknFWdySqRmie7bCM/edit?usp=sharing")

Progress,  w200, Updating...,Gathering Current Version From Server, Effectivity Macro Updater
Progress, 50

;Progress, off
Sleep, 200
;msgbox, checking for updates
;SplashTextOn , 100, 100, Serial Macros Updater, Checking for update.
while wb.busy
{
Sleep 100
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
Sleep 3000
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
Sleep 3000
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
;Sleep 100
;}


;MSGBOX, %clipboard%
settimer, versiontimeout, Off
		activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
Amonh /=2
amonw /=2

amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
Titletext := " Would you like to open the Cat Box site to download the latest version"


msgbox,262148,Effectivity Macro Updater, New update found. Would you like to open the Cat Box site to download the latest version?`n`nThe list below is what has changed.%whatisnew%
IfMsgBox, yes 
{
IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate
Run, https://cat.box.com/s/eghbsbas6d2qdwsy7y8cyfdglv2331mb
}
Else
Sleep 500
IniWrite, 14,  C:\Serialmacro\Settings.ini,update,updaterate	
IniWrite, %A_now%,  C:\Serialmacro\Settings.ini,  update,lastupdate

return
}
Gui,2:Destroy
return
}

versiontimeout:
{
Versioncount++
If Versioncount = 60
{
Progress,  w200,Updating..., Error Occured.Cannot connect to server for update check, please check for internet connection., Effectivity Macro Updater
Progress, 0
Sleep 3000
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