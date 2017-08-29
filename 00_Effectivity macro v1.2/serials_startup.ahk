

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


 
 

IfNotExist, C:\SerialMacro
{
FileCreateDir, C:\SerialMacro
sleep 100
}

IfNotExist, C:\SerialMacro\icons
{
FileCreateDir, C:\SerialMacro\icons
sleep 100
}
IfnotExist, C:\Serialmacro\Settings.ini
{
FileAppend,  C:\Serialmacro\Settings.ini
sleep 1000
IniWrite, 20,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate

IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
}

IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
If Refreshrate = Error
{
IniWrite, 20,  C:\Serialmacro\Settings.ini,refreshrate,refreshrate

IniRead, refreshrate, C:\Serialmacro\Settings.ini, refreshrate ,refreshrate
}

FileInstall, C:\ArbortextMacrosn\serial macro\How to use Effectivity Macro.pdf, C:\SerialMacro\How to use Effectivity Macro.pdf,1

FileInstall, C:\ArbortextMacrosn\icons\serial.ico, C:\SerialMacro\icons\serial.ico,1
FileInstall, C:\ArbortextMacrosn\icons\paused.ico, C:\SerialMacro\icons\paused.ico,1

IfNotExist, C:\SerialMacro\redimage.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\redimage.png, C:\SerialMacro\redimage.png,1
}

IfNotExist, C:\SerialMacro\plussign.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\plussign.png, C:\SerialMacro\plussign.png,1
}

IfNotExist, C:\SerialMacro\Activeplus.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\activeplus.png, C:\SerialMacro\activeplus.png,1
}

IfNotExist, C:\SerialMacro\orange button.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\orange button.png, C:\SerialMacro\orange button.png,1
}

IfNotExist, C:\SerialMacro\paused.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\paused.png, C:\SerialMacro\paused.png,1
}

IfNotExist, C:\SerialMacro\start.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\start.png, C:\SerialMacro\start.png,1
}

IfNotExist, C:\SerialMacro\Running.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\Running.png, C:\SerialMacro\Running.png,1
}

IfNotExist, C:\SerialMacro\Stopped.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\Stopped.png, C:\SerialMacro\Stopped.png,1
}

IfNotExist, C:\SerialMacro\background.png
{
FileInstall, C:\ArbortextMacrosn\serial macro\background.png, C:\SerialMacro\background.png,1
}

IfnotExist C:\SerialMacro\uredimage.png
{
FileNamered := "C:\SerialMacro\redimage.png"
}
Else 
{
FileNamered := "C:\SerialMacro\uredimage.png"
}

IfnotExist C:\SerialMacro\uplussign.png
{
FileNameSearch := "C:\SerialMacro\plussign.png"
}
Else
{
FileNameSearch := "C:\SerialMacro\uplussign.png"
}

IfnotExist C:\SerialMacro\uActiveplus.png
{
FileNameCheck := "C:\SerialMacro\Activeplus.png"
}
Else
{
FileNameCheck := "C:\SerialMacro\uActiveplus.png"
}

IfNotExist C:\SerialMacro\uorange button.png
{
FileNameButton := "C:\SerialMacro\orange button.png"
}
Else
{
FileNameButton := "C:\SerialMacro\uorange button.png"
}




