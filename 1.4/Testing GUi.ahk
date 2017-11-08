#SingleInstance force
Hide = 0
Gui, add, edit, x5 y5 h30 w200 vthree ,

Gui, add, edit, xp+205 h30 w200  vone
Gui, add, edit, xp+205 h30 w200   vtwo
;~ Gui 1:Add, Picture, xp yp w5 h5 +0x4000000  BackGroundTrans vPrefix_image, C:\SerialMacro\images\red_image.png
gui,add, Button, Default xp+205 gapply,Apply
gui,add, Button,  xp+100 ,Add
Gui 1:Add, Picture, x10 yp+50 w5 h5 +0x4000000  BackGroundTrans vtest_image, C:\SerialMacro\images\red_image.png
Gui 1:Add, Picture, xp+50 w5 h5 +0x4000000  BackGroundTrans vPrefix_image, C:\SerialMacro\images\red_image.png
Gui 1:Add, Picture, xp+50 w5 h5 +0x4000000  BackGroundTrans vENG_image, C:\SerialMacro\images\red_image.png
Gui, Add, Button, x30 Gtoggle_image_hide, Toggle image Off
Gui, Add, Button, xp+100 Gtoggle_image_show, Toggle image on
Gui, Add, Button, xp+100 Gone,One
Gui, Add, Button, xp+50 Gtwo, two
Gui, Add, Button, xp+50 Goff, Off
Gui, Show, w800, test gui
return

GuiClose:
ExitApp

one:
	GuiControl, Show, Prefix_image
	GuiControl, Hide, Eng_image
	GuiControl,Hide,test_image
	return

	two:
	{
	GuiControl, Show, Prefix_image
	GuiControl, Show, Eng_image
		return
	}


Off:
	GuiControl, Hide, Prefix_image
	GuiControl, Hide, Eng_image
	GuiControl,Hide,test_image
	return

;~ check:
;~ gui,Submit,NoHide
;~ GuiControlGet,one
;~ GuiControlGet,two
;~ GuiControlGet,three

;~ If three =
	;~ GuiControl, Show, Prefix_image
;~ else
	;~ GuiControl, Hide, Prefix_image

;~ If one =
	;~ GuiControl, Show, Eng_image
;~ else
	;~ GuiControl, Hide, Eng_image
;~ return

Apply:
If (Hide = "1")
GuiControl,Hide,test_image
GuiControl,,one,
GuiControl,,two,
GuiControl,,three,
;~ gosub check
Sleep 1000
If (!Hide)
GuiControl,Show,test_image

gui,Submit,NoHide
return

Toggle_image_hide:
{
GuiControl,Hide,test_image
Hide = 1
	return
}

Toggle_image_Show:
{
GuiControl,Show,test_image
Hide = 0
	return
}