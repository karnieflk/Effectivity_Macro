#Include Unit_testing\Yunit.ahk
#Include Unit_testing\Window.ahk
;~ #Include Unit_testing\StdOut.ahk
#Include Unit_testing\JUnit.ahk
#Include Unit_testing\OutputDebug.ahk

Yunit.Use(YunitWindow, YunitJUnit, YunitOutputDebug).Test(Folder_Function_Check, File_Function_Check,INI_File_Tests,temp_File_Functions_Check,Macro_Update_Check_Functions,Formatting_Functions_Check)

class Folder_Function_Check
{
    Folder_Create()
    {
    Result := Folder_Create("Macro_Test_Folder")
              Yunit.assert(Result == 0, "Error Creating Folder on Root C drive")

                Result := FileExist("C:\Macro_Test_Folder")
                 FileRemoveDir, C:\Macro_Test_Folder, 1
                Yunit.assert(Result !="" , "Did not create the folder")
                Result.Destroy
      }

Class Folder_Exists_Function
{
   Folder_Not_Exist_Check()
 {
Result := Folder_Exist_Check("Macro_Test_Folder_Not_Exist")
Yunit.assert(Result == "Macro_Test_Folder_Not_Exist - Folder_Not_Exist")
Result.Destroy
}

   Folder_Exists_Check()
    {
        FileCreateDir, C:\Macro_Test_Folder
        sleep 500
Result := Folder_Exist_Check("Macro_Test_Folder")
 FileRemoveDir, C:\Macro_Test_Folder, 1
         Yunit.assert(Result == "Macro_Test_Folder - Folder_Exist")
         Result.Destroy
  }

}}

class File_Function_Check
{
  File_Create()
{

     Result := File_Create("Macro_Testing.ini" , "1")
     Yunit.assert(Result == 0 , "Error Creating File")
          Result := FileExist(A_Desktop "\Macro_Testing.ini")
          FileDelete, %A_Desktop%\Macro_Testing.ini
                Yunit.assert(Result !="" , "Function not creating the file")
                Result.Destroy
}

Class File_Exist_Function
{

No_File_Exist_Check()
{
    Result := File_Exist_Check("Macro_Testing_not_exist.ini")
    Yunit.assert(Result == "Macro_Testing_not_exist.ini - File_Not_Exist")
    Result.Destroy
}



File_Exist_Check()
{
    FileAppend, test, %A_Desktop%\Macro_Testing.ini
      Result := File_Exist_Check("Macro_Testing.ini", "1")
     FileDelete, %A_Desktop%\Macro_Testing.ini
      Yunit.assert(Result == "Macro_Testing.ini - File_Exist")
      Result.Destroy
}}}

Class INI_File_Function_Tests
{

Ini_File_Read()
{

     Inifile := A_Desktop "\Macro_Testing.ini"
      FileAppend,,%inifile%
      sleep 500
Loop, 10
IniWrite,%A_index%,%inifile%,Testing,Test_Section%A_index%

    Load_ini_file(inifile)
    FileDelete, %inifile%
    Loop, 10
       Yunit.assert(Test_Section A_index == A_index)

inifile.Destroy
 }

Ini_File_Write()
{
       Inifile := A_Desktop "\Macro_Testing.ini"
  FileAppend,,%inifile%

Loop, 10
IniWrite,%A_index%,%inifile%,Testing,Test_Section%A_index%
    Load_ini_file(inifile)

  Write_ini_file(inifile)

    Load_ini_file(inifile)
           FileDelete %inifile%
    Loop, 10
       Yunit.assert(Test_Section A_index == A_index)

       inifile.Destroy
    }
}

Class temp_File_Functions_Check
{
    Temp_File_Delete()
{
     File_Location := A_Desktop
    File_Name := "Unit_test_Temp_File.txt"
    FileAppend, Test, %File_Location%\%File_name%
    sleep 500
Result := Temp_File_Delete(File_Location,File_Name)
Yunit.assert(Result == File_Name " - File Found and Deleted")
}

Class Temp_File_Read_Function
{
    Test_For_Failure_to_Read()
{

    File_Location := A_Desktop
    File_Name := "Unit_test_Temp_File_Failure_To_Read.txt"

Result := Temp_File_Read(File_Location,File_Name)
Yunit.assert(Result == "Null")
}

Test_For_Successful_Read()
{
    File_Location := A_Desktop
    File_Name := "Unit_test_Temp_File.txt"
    FileAppend, Test, %File_Location%\%File_name%

Result := Temp_File_Read(File_Location,File_Name)
FileDelete, %File_location%\%File_Name%
Yunit.assert(Result == "Test")
}}}

Class Macro_Update_Check_Functions
{
    class Calculations
    {
        Calculate_Days_Since_Last_Update()
{
    Updatestatus := A_Now
    Newtime = -10
    EnvAdd, UpdateStatus, %Newtime%, days

        Days := Calculate_Days_Since_Last_Update(updatestatus)
        Yunit.assert(Days == "10")
 }  }

                class Check_Doc_Title
                {

                 Test_For__Found_title()
                {
                    Gui 5:add, Text,, Test Window!
                    gui 5:Show,w500, Serial Version Unit Testing Gui
                    Correct_answer = Serial Version Unit Testing Gui

                   Result :=  Check_Doc_Title()

                       Gui 5:Destroy
                        Yunit.assert(Result == Correct_answer, "Could Not find ""Serial Version Unit Testing Gui"" Gui Window" )
            }

           Test_For__Not_Found_Title()
            {
                        Result :=  Check_Doc_Title()
                          Yunit.assert(Result == "Yunit Testing" , "Negative response check failed")
            }
                }

                class Format_Serial_check_title
                {
                    Test_For__Found_title()
                    {
                    Title := "Serial Version #99"
                Result :=   Format_Serial_Check_Title(Title)
                Yunit.assert(Result == "99" , "Did not format title correctly")
                }

                TEst_for__Not_found_title()
                    {
                    Title := "Not_Found"
                Result :=   Format_Serial_Check_Title(Title)
                Yunit.assert(Result == "Not_Found" , "Did not Pass Not_found")
                }

                }


}


    class Formatting_Functions_Check
    {

        Combinecount()
        {
            Prefix_Store_Array := "one,two,six,ten,pat,sam,big,fog,dog,log,mog,mug,cat,"
            Correct_answer := "13"
        Result := Combinecount(Prefix_Store_Array)
        Yunit.assert(Result == Correct_answer, "Did not add digits")
    }
             Add_Digits()
             {
                Serial_number = 1-5`)
                Correct_answer = 00001-00005`)
                    Result:= Add_Digits(Serial_Number)
                        Yunit.assert(Result == Correct_answer, "Did not add digits")

              Serial_number = 10000-500000`)
                Correct_answer = 10000-500000`)
                    Result:= Add_Digits(Serial_Number)
                    Yunit.assert(Result == Correct_answer, "Did not let variable pass through")

              Serial_number = 500-550`)
                Correct_answer = 00500-00550`)
                    Result:= Add_Digits(Serial_Number)
                    Yunit.assert(Result == Correct_answer, "Did not let variable pass through")
                }

Put_Formatted_Serials_into_Array()
{
 Test_Serials := "TDK00001,`nTDK00002,`nTDK00003,`n"
Result_array := Object()
Result_array := Put_Formatted_Serials_into_Array(Test_Serials)
For Index, Element in Result_array
  Yunit.assert(Element == "TDK0000" Index ","," Failed on Index" index  )

}

Extract_Serial_Array()
{
       Test_array := Object()
       Loop, 5
        Test_array.Insert("TRD0000"A_Index "-0000"A_Index)

      Correct_answer =
      (
TRD00001-00001,
TRD00002-00002,
TRD00003-00003,
TRD00004-00004,
TRD00005-00005,`n
)
 Result := Extract_Serial_Array(Test_Array)
 ;~ MsgBox, % result
 Yunit.assert(REsult == Correct_answer )
}

Formatted_Text_Serial_Count()
{
Test_Serials =
(
TRD00001-00001,
TRD00002-00002,
TRD00003-00003,
TRD00004-00004,
TRD00005-00005,`n
)
Correct_answer := 5
 Result := Formatted_Text_Serial_Count(Test_Serials)
  Yunit.assert(REsult == Correct_answer )
}

One_Up_all()
{
    Serial_Store_Array := Object(), Result_array := Object()

    Loop, 5
        Serial_Store_Array.Insert("TR" A_Index "0000"A_Index "-00005")

   Result_array :=  One_Up_All(Serial_Store_Array)
   For index, Element in Result_array
{
      Yunit.assert(Element  == "TR" A_Index "00001-99999" )
}

}

class Copy_Selected_Text_Function
{


Test_Is_Selected_Test()
{
Correct_answer = Test Window`,
                 Gui,6:add, Edit,,Test Window
                    gui 6:Show,w500, Copy Test
                    WinActivate Copy test
                REsult :=  Copy_selected_Text()
                  Gui, 6:Destroy
                  Yunit.assert(REsult == Correct_answer)

 }

 Text_Is_Not_Selected_Test()
 {

    Correct_answer = No_Text_Selected

                 Gui,6:add, Edit,,
                    gui 6:Show,w500, Copy Test
                    WinActivate Copy test
                REsult :=  Copy_selected_Text()
                 Gui, 6:Destroy
                  Yunit.assert(REsult == Correct_answer, "Result should be No_Text_Selected" )

}}

class Check_For_Single_Serials
{
            Text_Is_All_Single_Serials()
            {
                PreFormatted_Text := "TRD00001,HAT00001,FRT00001,LFT00001,RHT00001"

                Correct_answer := "TRD00001-00001,HAT00001-00001,FRT00001-00001,LFT00001-00001,RHT00001-00001,"
            Result :=  Check_For_Single_Serials(PreFormatted_Text)
                      Yunit.assert(REsult == Correct_answer )
    }

            Some_Text_Is_Single_Serials()
            {
                PreFormatted_Text := "TRD00001-00005,HAT00001,FRT00001,LFT00001-10002,RHT00001"

                Correct_answer := "TRD00001-00005,HAT00001-00001,FRT00001-00001,LFT00001-10002,RHT00001-00001,"
            Result :=  Check_For_Single_Serials(PreFormatted_Text)
               Yunit.assert(REsult == Correct_answer )
    }

                Text_no_Single_Serials()
            {
                PreFormatted_Text := "TRD00001-00005,HAT00001-00001,FRT00001-00001,LFT00001-10002,RHT00001-00001,"
            Result :=  Check_For_Single_Serials(PreFormatted_Text)
                   Yunit.assert(REsult == PreFormatted_Text )
    }
}


}



