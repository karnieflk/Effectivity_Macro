#Include Unit_testing\Yunit.ahk
#Include Unit_testing\Window.ahk
;~ #Include Unit_testing\StdOut.ahk
#Include Unit_testing\JUnit.ahk
#Include Unit_testing\OutputDebug.ahk

Yunit.Use(YunitWindow, YunitJUnit, YunitOutputDebug).Test(Folder_Function_Check, File_Function_Check)

class Folder_Function_Check
{

   _Folder_Exists__Test_FolderNotExist() ; Needs to have the "_" in front so it runs first
    {
Result := Folder_Exist_Check("Macro_Test_Folder")
Yunit.assert(Result == "Macro_Test_Folder - Folder_Not_Exist", "Folder Exists - Delete folder on Root C Drive and Rerun tests")
}


    Folder_Create()
    {
    Result := Folder_Create("Macro_Test_Folder")
              Yunit.assert(Result == 0, "Error Creating Folder on Root C drive")


                Result := FileExist("C:\Macro_Test_Folder")
                Yunit.assert(Result !="" , "Cannot find Folder on Root C drive")
      }

    Folder_Exists__Test_FolderExists()
    {
Result := Folder_Exist_Check("Macro_Test_Folder")
 FileRemoveDir, C:\Macro_Test_Folder, 1
         Yunit.assert(Result == "Macro_Test_Folder - Folder_Exist")
  }

}

class File_Function_Check
{
    Begin()
    {
     }

_File_Exist_Check__Test_NoFileExist() ; Needs to have the "_" in front so it runs first
{
    Result := File_Exist_Check("Macro_Testing.ini")
    Yunit.assert(Result == "Macro_Testing.ini - File_Not_Exist", "File exists, delete file in root C folder")
}

File_Create()
{

     Result := File_Create("Macro_Testing.ini" , "1")
     Yunit.assert(Result == 0 , "Error Creating File")
          Result := FileExist(A_Desktop "\Macro_Testing.ini")
                Yunit.assert(Result !="" , "Function not creating the file")
}

File_Exist_Check__Test_FileExist()
{
      Result := File_Exist_Check("Macro_Testing.ini", "1")
        FileDelete, %A_Desktop%\Macro_Testing.ini
    Yunit.assert(Result == "Macro_Testing.ini - File_Exist")
}





    End()
    {
    }
}