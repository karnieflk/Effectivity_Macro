aboutmacro:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
			  
   ;msgbox, Prefix is %Guitextlocation%`n Serial is %Second_Effectivity_Numbers%-%Second_Effectivity_Numbers%`n Mod is %Modifier%
   Gui, 7:Add, Picture, x0 y0 w525 h300 +0x4000000, C:\SerialMacro\background.png
   gui, 7:add, Text, xp+5 yp+5 w500 h40 BackgroundTrans, This program was designed to help increase the speed and accuracy of entering machine Effectivity prefixes form the Pubtool into ACM.
   gui, 7:add, Text, xp yp+45 w500 h20 BackgroundTrans, To get the latest version, go to the File menu and select Check for updates.
   gui, 7:add, Text, xp yp+25 w300 h20 BackgroundTrans,  The location of the macro is at the following box account:
   gui 7:font, CBlue Underline
   gui, 7:add, Text, xp+275 yp w500 h20 BackgroundTrans gboxlink, https://cat.box.com/s/eghbsbas6d2qdwsy7y8cyfdglv2331mb
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

emaillink:
{
   Run,  mailto:Karnia_Jarett_S@cat?Subject=Effectivity Macro
   return
}

boxlink:
{
   Run, https://cat.box.com/s/eghbsbas6d2qdwsy7y8cyfdglv2331mb
   return
}

SerialbreakquestionGUI:
{
   activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
   
   Amonh /=2
   amonw /=2
   
   amonx := amonx + (amonw/2)
   amony := amony + (amonh/2)
   gui 1: -alwaysontop
   Gui, 8:Add, Picture, x0 y0 w400 h90 +0x4000000, C:\SerialMacro\background.png
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
   Oneupserial = Yes
   Gui, 8:destroy
   combineser = 1
   GOsub, radio1h
   Return
}

combinequstion:
{
   UnPausescript()
   GuiControlGet, checked,, createexcel
   gui 1: +alwaysontop
   Gui, 8:destroy
   combineser = 1
   GOsub, radio1h
   Return
}

keepseperated:
{
   UnPausescript()
   GuiControlGet, checked,, createexcel
   gui 1: +alwaysontop
   Gui, 8:destroy
   combineser = 0
   GOsub, radio1h
   Return
}

SerialsGUIscreen:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
   
   Menu, BBBB, Add, &Check For Update , Versioncheck
   Menu, BBBB, Add, &Options, OptionsGui
   Menu, BBBB, Add, 
   Menu, CCCC, Add, &Run							(Crtl + 2), Enterallserials
   Menu, CCCC, Add, &Pause/Unpause 				(Pause / Insert), pausesub
   Menu, CCCC, Add, &Stop Macro					(ESC), Exitprogram
   Menu, CCCC, Add, &Reload Macro, restartmacro
   Menu, CCCC, Add, &Reload Macro with Current Effectivity, restartmacroEffectivity
   Menu, BBBB, Add, &Exit							(Ctrl + Q), Quitapp
   Menu, DDDD, Add, &How To Use					(F1), HowTo
   Menu, DDDD, Add, &About , Aboutmacro
   
   Menu, MyMenuBar, Add, &File, :BBBB
   Menu, MyMenuBar, Add, &Macro, :CCCC
   Menu, MyMenuBar, Add, &Help, :DDDD
   SplashTextOff
   If totalprefixes < 1
   {
      TotalPrefixzero = 0
   }
   Else 
      TotalPrefixzero = %totalprefixes%
   
   gui 1:add, Edit, x10 y50 w390 h240  vEditField,%editfield%
   
   gui 1:add, Edit, xp yp w390 h240 vEditField2,%editfield2%
   
   Gui 1:Add, Picture, x315 y310 w50 h50 +0x4000000  BackGroundTrans vStarting gstartmacro , C:\SerialMacro\Start.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vRunning, C:\SerialMacro\Running.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vpaused  gpausesub, C:\SerialMacro\Paused.png
   Gui 1:Add, Picture, xp yp w50 h50 +0x4000000 BackGroundTrans  vStopped grestartmacro, C:\SerialMacro\Stopped.png
   Gui, 1:Add, Picture, x0 y0 w410 h400 +0x4000000 , C:\SerialMacro\background.png
   
   Gui 1:Add, Edit, xp+165 yp+343 w110 h20  vnextserialtoadd, %nextserialtoaddv%
   
   Gui 1:Add, Text, x5 y5 w300 h25 BackgroundTrans +Center vreloadprefixtext, There are a total of %totalprefixzero% Effectivity to add to ACM
   
   Gui 1:add, Radio, xp+25 yp+25 w130 h20 BackGroundTrans vradio1 gradio1h, Effectivity to be added: 
   
   
   
   Gui 1:add, Radio, xp+155 yp w140 h20 BackGroundTrans vradio2 gradio2h, Effectivity already added
   Gui 1:Add, Text, xp-170 Yp+265 W250 h13 BackGroundTrans vserialsentered, Number of Effectivity successfully added to ACM = %Serialcount%
   
   Gui 1:Add, Text, Xp Yp+15 w250 h13  BackGroundTrans , If macro is operating incorrectly, press Esc to reload
   Gui 1:Add, Text, xp yp+15 w250 h13  BackGroundTrans , Or press Pause Button on keyboard to Pause macro, Press Pause again to resume macro.
   Gui 1:Add, Text, xp yp+20 w145 h20  BackgroundTrans , Next Effectivity to add to ACM:
   
   Gui 1:Menu, MyMenuBar
   Gui 1:Show,  x%amonx% y%amony% , %Effectivity_Macro%
   gui 1: +alwaysontop
   Guicontrol,hide, Editfield2,
   Guicontrol,show, Editfield,
   Guicontrol, Focus, Editfield
   IfExist  C:\SerialMacro\Tempcount.txt
   {
      FileRead, Serialcount,C:\SerialMacro\Tempcount.txt
      GuiControl,1:,serialsentered, Number of Serials successfully added to ACM = %Serialcount%	
      FileDelete, C:\SerialMacro\Tempcount.txt
   }
   Gui 1:Submit, NoHide
   Return
}

OptionsGui:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   Iniread, refreshrate,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
   Iniread, imagesearchoption,  C:\Serialmacro\Settings.ini,searches,imagesearchoption
   gui 10: +alwaysontop
   Gui , 1: -AlwaysOnTop
   ;gui 10:add,radio, checked%mouseclicks% x0 y0 w250 h20 vmouseposclicks gmouseclickss, Mouse click method 
   ;gui 10:add,radio, checked%imagesearchoption% xp yp+25 w317 h20 vimagesearchoption gimagessearching, Image search method `(Not fully functional`)
   gui 10:add, text, x5 y5 w320 h20 ,Refreash ACM Rate (After how many entered effectivity)
   gui 10:add, edit, xp+275 yp-3 w30 veditfield5 , %refreshrate%
   gui 10:add, button, xp-251 yp+26 h20 w75 Default gsavesets, Save Settings
   Gui, 10:Add, Picture, x0 y0 w325 h100 +0x4000000 , C:\SerialMacro\background.png
   gui 10:show, x%amonx% y%amony% w325 h100
   Guicontrol,10:, editfield5, %refreshrate%
   gui, 10:submit, nohide
   return
}

savesets:
{
   Gui , 1: +AlwaysOnTop
   ;msgbox, %refreshrate%
   GuiControlGet,Refreshrate,,editfield5
   GuiControlGet,imagesearchoption
   GuiControlGet,mouseposclicks
   
   If imagesearchoption = 1
   Entermethod = 1
   
   If mouseposclicks = 1
   Entermethod = 2
   
   IniWrite, %refreshrate%,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate
   Iniwrite, %Entermethod%,  C:\Serialmacro\Settings.ini,searches,entermethod
   Iniwrite, %mouseposclicks%,  C:\Serialmacro\Settings.ini,searches,mouseposclicks
   Iniwrite, %imagesearchoption%,  C:\Serialmacro\Settings.ini,searches,imagesearchoption
   gui 10:submit, nohide
   gui 10:destroy
   return
}

10guiclose:
{
   Gui , 1: +AlwaysOnTop
   gui 10:destroy
   return
}


mouseclickss:
{
   Entermethod = 2
   Iniwrite, %Entermethod%,  C:\Serialmacro\Settings.ini,searches,entermethod
   return
}

imagessearching:
{
   Entermethod = 1
   Iniwrite, %Entermethod%,  C:\Serialmacro\Settings.ini,searches,entermethod
   return
}


startmacro:
{
   skipbox = 0
   gosub, enterallserials
   return
}

rerunclicks:
{
   Stop = 1
   Send {shift down}{Shift up}
   
   Return
}

pauses:
{
   UnPausescript()
   return
}

radio1h:
{
   GuiControl,, radio2, 0
   GuiControl,, radio1,1
   ;GuiControl,, radio3,0
   ;Storeeditfield2 = %editfield2%`
   Guicontrol,show, Editfield,
   Guicontrol,hide, Editfield2,
   ;Guicontrol,hide, Editfield3,
   gui, submit, nohide
   
   return
}


radio2h:
{
   GuiControl,, radio1,0
   GuiControl,, radio2,1
   ;GuiControl,, radio3,0
   ;Storeeditfield1 = %editfield%
   Guicontrol,show, Editfield2,
   Guicontrol,hide, Editfield,
   ;Guicontrol,hide, Editfield3,
   
   gui, submit, nohide
   ;Guicontrol,, Editfield, %EditField2%
   return
}

radio3h:
{
   GuiControl,, radio1,0
   GuiControl,, radio2,0
   ;GuiControl,, radio3,1
   ;Editfieldtemp = %Editfield%
   ;Editfield = %EditfieldCombine%
   ;Guicontrol,,Editfield,%EditfieldCombine%
   ;Gui, 1:Submit, nohide
   ;Guicontrol,show, Editfield3,
   Guicontrol,hide, Editfield,
   Guicontrol,hide, Editfield2,
   
   ;Guicontrol,, Editfield, %EditField2%
   return
}

Guiclose:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   Titletext := "Are you sure you want to quit"
   Settimer, winmovemsgbox, 20
   
   msgbox,262148,Quit %Effectivity_Macro%, Are you sure you want to quit?
   ifMsgBox Yes
   {
      ;gui destroy
      ExitApp
   }else  {
      Return
   }
   Return
}

restartmacro:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   Titletext := " Are you sure that you want to reload the program"
   Settimer, winmovemsgbox, 20
   msgbox,262148,Reload %Effectivity_Macro%, Are you sure that you want to reload the program?
   IfMsgBox yes
   {
      Reload
   }else  {
      return
   }
   return
}


restartmacroEffectivity:
{
   activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   Titletext := " Are you sure that you want to reload the program"
   Settimer, winmovemsgbox, 20
   msgbox,262148,Reload %Effectivity_Macro%, Are you sure that you want to reload the program?
   IfMsgBox yes
   {
      GuiControlGet, nextserialtoadd
      GuiControlGet, EditField
      GuiControlGet, EditField2
      TempSavefile = %nextserialtoadd%`,`n%editfield%
      
      FileAppend, %TempSavefile%, C:\SerialMacro\TempAdd.txt 
      FileAppend, %EditField2%, C:\SerialMacro\TempAdded.txt 
      FileAppend, %totalprefixes%, C:\SerialMacro\Tempamount.txt
      FileAppend, %Serialcount%, C:\SerialMacro\Tempcount.txt
      Reload
   }else  {
      return
   }
   return
}


Macrotimedout:
{
   IniRead, amonx,  C:\Serialmacro\Settings.ini,Timeoutwindow,Xposition
   IniRead, amony,  C:\Serialmacro\Settings.ini,Timeoutwindow,Yposition
   
   If amonx = Error
   {
      activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   }
   
   If Amony = Error
   {
      activeMonitorInfo( amonx,Amony,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.

   }
   
   ;msgbox, Prefix is %Guitextlocation%`n Serial is %Second_Effectivity_Numbers%-%Second_Effectivity_Numbers%`n Mod is %Modifier%
   Gui, 3:Add, Picture, x0 y0 w385 h215 +0x4000000, C:\SerialMacro\background.png
   gui, -alwaysontop
   
   
   gui, 3:add, Text, x10 y10 w360  h20 BackgroundTrans,  The delay is most likely due to one of the following:
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - The Eng Model needs to be manually added with the drop down menu.
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - Effectivity # Does not Exist in ACM
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - ACM is running slow
   gui, 3:add, Text, xp yp+21 w360  h20 BackgroundTrans,  - Something bad happened and the effectivity page is no more.
   
   gui, 3:add, Text, xp yp+21 w360  h40 BackgroundTrans,  Please press the one of the buttons below after you Idenitify the cause. This action will reflect in the Effectivity Added Screen.
   gui, 3:add, Button, xp+1 yp+45 w80 h50 BackgroundTrans gEngmodel,  More than one Engineering Model
   gui, 3:add, Button, xp+90 yp w80 h50 BackgroundTrans gSerialnogo,  Effectivity Does Not Exist
   gui, 3:add, Button, xp+90 yp w80 h50 BackgroundTrans gacmlong,  ACM Took to long
   gui, 3:add, Button, xp+90 yp w80 h50 BackgroundTrans  gAnarchy,  Something else happened :(
   Gui, 3:Show, x%amonx% y%amony%, Macro Timed Out %Effectivity_Macro%
   Guicontrol,hide, Start
   Guicontrol,show, paused
   Guicontrol,hide, Stopped
   Guicontrol,hide, Running
   Gui, Submit, NoHide
   gui, 3: +alwaysontop
   Pausescript()
   return
}






Howto:
{
   splashtexton,,Effectivity Macro, Loading PDF
   Run, C:\SerialMacro\How to use Effectivity Macro.pdf
   sleep 2000
   SplashTextOff
   return
}              