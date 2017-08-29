Coordmode, mouse, screen
/*
insert::
{
Pause
return
}

esc::
{
reload
}

$#z::
{
Hotkey, $+Lbutton, Grabscreen,  on
Hotkey, $+Rbutton, Grabscreen, on
Gosub, Loadimagefiles
Gosub, Addsearch
Gosub, Prefixboxsearch
return
}
*/

Loadimagefiles:
{
msgbox, loading images
Hotkey, $+Lbutton,  Off
Hotkey, $+Rbutton,  Off
If pToken = 0
{
pToken := Gdip_Startup()
}

SetKeyDelay, -1
SetMouseDelay, -1
SetDefaultMouseSpeed, 0
SetWinDelay, -1
SetBatchLines -1
ox = ""
oy= ""
needlecountadd = 0
needlecountprefix = 0
CurrMonPrefix := ""
FileCount := ""
+++SetTitleMatchMode, 2
Sleep 200
mousemove 100,100
Gosub, Filecounts
BreakLoop := 0
sleep 100
wingetTitle, title, a
WinActivate, %title%
;Gosub, SerialFullScreen
return
}

addsearchrefresh:
{
WinActivate %title%
mousemove 100, 100
Counter = 0
CurrMonPrefix := GetCurrentMonitorcpn()
If pToken = 0
{
pToken := Gdip_Startup()
}
If filecountadd = 0
{
;Gosub, ArborFullScreenCPN
Gosub, GUImessageboxaddbutton
}
Loop, %needlecountadd%
{
Counter++
bmpHaystack := Gdip_BitmapFromScreen(CurrMonPrefix)
RETSearch := Gdip_ImageSearch(bmpHaystack,FoundAddbutton,List,0,0,0,0,0,0,0,0)
If RETSearch > 0
{
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
gosub, addsearchrefresh
gosub, EnterCPn
break
}
}
If RETSearch = 0
{

;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
Return
}
Running_script = 0
Return
}

filecounts:
{
needlecountprefix = 0
needlecountadd = 0
SplashTextOn,,20,Serial Macro, Loading Image files
Filecountprefix := ComObjCreate("Shell.Application").NameSpace("C:\SerialMacro\PrefixImages").Items.Count
Filecountadd := ComObjCreate("Shell.Application").NameSpace("C:\SerialMacro\addbuttonImages").Items.Count
;Filecountend = %Filecountprefix%

Loop, %Filecountadd%
{
needlecountadd++

FileNameaddbutton := "C:\SerialMacro\AddbuttonImages\User_" needlecountadd ".png"
bmpNeedleaddbutton%needlecountadd% := Gdip_CreateBitmapFromFile(FileNameaddbutton)
sleep 50
}
Needlecountadd = 0
Loop, %FileCountprefix%
{
needlecountprefix++
FileNameprefix := "C:\SerialMacro\PrefixImages\User_" needlecountprefix ".png"
bmpNeedleprefix%needlecountprefix% := Gdip_CreateBitmapFromFile(FileNameprefix)
sleep 50
}

splashtextoff
return
}

Addsearch:
{
counter = 0
msgbox, addsearch
sleep 500
WinActivate %title%
mousemove 100, 100
Counter = 0
CurrMonPrefix := GetCurrentMonitorcpn()
If pToken = 0
{
pToken := Gdip_Startup()
}
If filecountadd = 0
{
;Gosub, ArborFullScreenCPN
Gosub, GUImessageboxaddbutton
sleep 3000
gosub, filecounts
gosub, addsearch
}
Loop, %needlecountadd%
{
Counter++
bmpHaystack := Gdip_BitmapFromScreen(CurrMonPrefix)
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedleaddbutton%counter%,List,0,0,0,0,0,0,0,0)
If RETSearch > 0
{
msgbox,found addsearch
FoundAddbutton = bmpNeedleaddbutton%counter%
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
gosub, EnterCPn
break
}
}
If RETSearch = 0
{
msgbox, not found addsearch
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
Gosub, GUImessageboxaddbutton
}
msgbox, addsearch returning
Running_script = 0
Return
}

Prefixboxsearch:
{
counter = 0
msgbox, search prefix
WinActivate %title%
mousemove 100, 100
Counter = 0
CurrMonPrefix := GetCurrentMonitorcpn()
If pToken = 0
{
pToken := Gdip_Startup()
}
If filecountprefix = 0
{
;Gosub, ArborFullScreenCPN
Gosub, GUImessageboxprefix
gosub, filecounts
}
Loop, %needlecountprefix%
{
Counter++
bmpHaystack := Gdip_BitmapFromScreen(CurrMonPrefix)
RETSearch := Gdip_ImageSearch(bmpHaystack,bmpNeedleaddbutton%counter%,List,0,0,0,0,0,0,0,0)

If RETSearch > 0
{
msgbox, found prefixbox
FoundAddbutton = bmpNeedleaddbutton%counter%
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
gosub, EnterCPn
break
}
}
If RETSearch = 0
{
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
Gosub, GUImessageboxprefix
Return
}
;Gosub, finished
Running_script = 0
Return
}

EnterCPn:
{
Running_script = 1
;CoordMode, mouse, Relative
Loop, Parse, LIST, `n
{
WinActivate, %title%
WinWaitActive, %title%,,3
StringSplit, Coord, A_LoopField, `,
winactivate, %title%
WinWaitActive, %title%,,3
coord1 = %coord1% - 20
coord2 = %coord2% - 20
MouseMove, %Coord1%, %Coord2%, 0
Click
msgbox, clicked
}
Running_script = 0
}

finished:
{
Loop, %counter%
{
Gdip_DisposeImage(bmpNeedle%counter%)
counter--
}
Return
}

GUImessageboxaddbutton:
{
WinGetPos,px,py,ph,pw, %title%
ph /= 4
pw /= 3
px += %pw%
py += %ph%
Gui, 5:+AlwaysOnTop
Gui, 5:Color, White
Gui, 5:Add, Text, x160 y8 w419 h100, Unable to find the add box. Please click on the Add button, then Hold down shift and click and drag the mouse over the add box after you click OK in this window.
Gui, 5:Add, Button, x480 y168 w144 h51 gButtonCPNtags +Default, Ok
Gui, 5:Add, Button, x10 y168 w144 h51 gButtonCancel, Cancel
Gui , 5:show, w640 h227 , Oops!!!
WinMove Oops!!!,,px,py
pause, on
return
}

GUImessageboxprefix:
{
WinGetPos,px,py,ph,pw, %title%
ph /= 4
pw /= 3
px += %pw%
py += %ph%
Gui, 6:+AlwaysOnTop
Gui, 6:Color, White
Gui, 6:Add, Text, x160 y8 w419 h100, Unable to find the prefix. Please hold down shift and click and drag the mouse over the prefix box after you click OK in this window. Make sure to include the Prefix text above the box or program will have eratic behaviour. 
Gui, 6:Add, Button, x480 y168 w144 h51 gButtonprefixselect +Default, Ok
Gui, 6:Add, Button, x10 y168 w144 h51 gButtonCancel, Cancel
Gui , 6:show, w640 h227 , Oops!!!
WinMove Oops!!!,,px,py
pause on
return
}

6guiclose:
{
gui, 6:destroy
reload
return
}

5guiclose:
{
gui, 5:destroy
reload
return
}

ButtonCancel:
{
Gui, 5:Destroy
gui,6:destroy
reload
}

ButtonCPNtags:
{
pause, off
grabvar = 2
Gui, 5:Destroy
Gosub, SelectCPNTag
return
}


Buttonprefixselect:
{
pause, off
grabvar = 1
Gui, 6:Destroy
Gosub, SelectCPNTag
return
}


Grabscreen:
{
If grabvar = 1
{
SCW_ScreenClip2Win(clip :=0,email :=1 ,countLoops = %countLoop%)
}
If grabvar = 2
{
SCW_ScreenClip2Win(clip :=0,email :=2 ,countLoops = %countLoop%)
}
Return
}

SelectCPNTag:
{
Running_script = 1
WinActivate, %title%
WinWaitActive, %title%,,3
Mousemove 500,500
sleep 50
CurrMonPrefix := GetCurrentMonitorcpn()
countLoop := 1
;Gosub, loopFileName
Hotkey, $+Lbutton,Grabscreen, on
Hotkey, $+Rbutton,Grabscreen ,  on
Keywait,Shift, D,t10
Keywait, Shift, U,t10
Return
}


Mouselock:
{
mousemove Mox,Moy
Return
}

Capture:
{
Running_script = 1
monitorar := GetCurrentMonitorcpn()
activeMonitorInfo( amonX,AmonY,AmonW,AmonH,mx,my )
;Gosub,take_snapshot
Running_script = 0
return
}

loopFileName:
{
Loop
{
countLoopString := countLoop
newFileName := "C:\SerialMacro\PrefixImages\User_" countLoopString ".png"
IfExist, % newFileName
{
countLoop++
Goto, loopFileName
}
IfNotExist, % newFileName
{
Break
}
Return
}

take_snapshot:
{
Running_script = 1
if monitorar = 1
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(1,raster)
pBitmap2 := Gdip_CreateBitmap(70,16)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
Else if monitorar = 2
{
raster:=0x40000000 + 0x00CC0020
If pToken = 0
{
pToken := Gdip_Startup()
}
pBitmap := Gdip_BitmapFromScreen(2,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
Else if monitorar = 3
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(3,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
else if monitorar = 4
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(4,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
else if monitorar = 5
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(5,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
else if monitorar = 6
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(6,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
else if monitorar = 7
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(7,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
else if monitorar = 8
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(8,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
else if monitorar = 9
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(9,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
else if monitorar = 10
{
If pToken = 0
{
pToken := Gdip_Startup()
}
raster:=0x40000000 + 0x00CC0020
pBitmap := Gdip_BitmapFromScreen(10,raster)
pBitmap2 := Gdip_CreateBitmap(50, 50)
G2 := Gdip_GraphicsFromImage(pBitmap2), Gdip_SetSmoothingMode(G2, 4), Gdip_SetInterpolationMode(G2, 7)
Gdip_DrawImage(G2, pBitmap, 0, 0, 50, 50, mx, my, 50, 50)
Gdip_SaveBitmapToFile(pBitmap2, newFileName)
Gdip_DeleteGraphics(G2), Gdip_DisposeImage(pBitmap), Gdip_DisposeImage(pBitmap2)
;If pToken !=0
;{
;Gdip_Shutdown(pToken)
;}
}
Running_script = 0
return
}

Windowshow:
{
settimer, watchcursor,  5
Gui, 7:+AlwaysOnTop
gui,7:-Caption +ToolWindow +AlwaysOnTop +LastFound
gui ,7:margin, 0, 0
Gui,7:Add, Picture, x0 y0 w50 h50, C:\ArbortextMacros\ProgramImages\Box.png
Gui,7:Show, w50 h50, movinggui
Gui,7:Color, FFFFFF
WinSet, TransColor, FFFFFF
return
}

WatchCursor:
{
CoordMode, mouse, Screen
MouseGetPos, oX, oY
op = %ox% - 35
ob = %oy% - 35
WinMove, movinggui,,op,ob
return
}

ShowCursor:
{
SystemCursor("On")
Exit
Return
}

SystemCursor(OnOff=1)
{
static AndMask, XorMask, $, h_cursor
,c0,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13
, b1,b2,b3,b4,b5,b6,b7,b8,b9,b10,b11,b12,b13
, h1,h2,h3,h4,h5,h6,h7,h8,h9,h10,h11,h12,h13
if (OnOff = "Init" or OnOff = "I" or $ = "")
{
$ = h
VarSetCapacity( h_cursor,4444, 1 )
VarSetCapacity( AndMask, 32*4, 0xFF )
VarSetCapacity( XorMask, 32*4, 0 )
system_cursors = 32512,32513,32514,32515,32516,32642,32643,32644,32645,32646,32648,32649,32650
StringSplit c, system_cursors, `,
Loop %c0%
{
h_cursor   := DllCall( "LoadCursor", "Ptr",0, "Ptr",c%A_Index% )
h%A_Index% := DllCall( "CopyImage", "Ptr",h_cursor, "UInt",2, "Int",0, "Int",0, "UInt",0 )
b%A_Index% := DllCall( "CreateCursor", "Ptr",0, "Int",0, "Int",0
, "Int",32, "Int",32, "Ptr",&AndMask, "Ptr",&XorMask )
}
}
if (OnOff = 0 or OnOff = "Off" or $ = "h" and (OnOff < 0 or OnOff = "Toggle" or OnOff = "T"))
$ = b
else
$ = h
Loop %c0%
{
h_cursor := DllCall( "CopyImage", "Ptr",%$%%A_Index%, "UInt",2, "Int",0, "Int",0, "UInt",0 )
DllCall( "SetSystemCursor", "Ptr",h_cursor, "UInt",c%A_Index% )
}
}

#IfWinActive
ArborFullScreencpn:
{
Running_script = 1
wingetTitle,Titlearbor,A
WinGetPos, Xarbor,yarbor,warbor,harbor, %titlearbor%
CurrMonPrefix := GetCurrentMonitorcpn()
SysGet,Aarea,MonitorWorkArea,%CurrMonPrefix%
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
MouseMove, mmx, mmy
Running_script = 0
return
}






GetCurrentMonitorcpn()
{
WinGetTitle, title,a
SysGet, numberOfMonitors, MonitorCount
WinGetPos, winX, winY, winWidth, winHeight, %title%
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


CoordMode, Mouse, Screen
GetCurrentMonitorst()
{
WinGetTitle, title,a
SysGet, numberOfMonitors, MonitorCount
WinGetPos, winX, winY, winWidth, winHeight, %title%
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

activeMonitorInfost( ByRef aX, ByRef aY, ByRef aWidth,  ByRef  aHeight, ByRef mouseX, ByRef mouseY  )
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
RemoveToolTip:
{
SetTimer, RemoveToolTip, Off
Tooltip
Return
}
Monitorareas:
{
msgbox, monitor areas
SysGet, MonitorCount, 80
SysGet, MonitorPrimary, MonitorPrimary
Loop %MonitorCount%
{  SysGet, Pos, MonitorWorkArea, %A_Index%
MonitorWorkArea[%A_Index%] := {l:PosLeft,r:PosRight,t:PosTop,b:PosBottom}
}
Msgbox, %MonitorCount% monitors
Return
}

#include <screen reduced>