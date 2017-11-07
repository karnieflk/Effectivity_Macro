#SingleInstance force
Gui, add, edit, h30 w200 vthree,
Gui, add, edit, xp+205 h30 w200 vone
Gui, add, edit, xp+205 h30 w200 vtwo
gui,add, Button, Default xp+205 gapply,Apply
gui,add, Button,  xp+100 ,Add
Gui 1:Add, Picture, x10 yp+50 w5 h5 +0x4000000  BackGroundTrans vtest_image, C:\SerialMacro\images\red_image.png
Gui, Add, Button, x30 Gtoggle_image_hide, Toggle image Off
Gui, Add, Button, xp+100 Gtoggle_image_show, Toggle image on
Gui, Show, w800, test gui
return

GuiClose:
ExitApp


Apply:
gui,Submit,NoHide
GuiControl,Hide,test_image
GuiControl,,one,
GuiControl,,two,
GuiControl,,three,
Sleep 1000
GuiControl,Show,test_image
return

Toggle_image_hide:
{
GuiControl,Hide,test_image
	return
}

Toggle_image_Show:
{
GuiControl,Show,test_image
	return
}