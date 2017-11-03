#SingleInstance force
Gui, add, edit, h30 w200 vthree,
Gui, add, edit, xp+205 h30 w200 vone
Gui, add, edit, xp+205 h30 w200 vtwo
gui,add, Button, Default xp+205 gapply,Apply
Gui, Show, w800, test gui
return

GuiClose:
ExitApp


Apply:
GuiControl,,one,
GuiControl,,two,
GuiControl,,three,
return