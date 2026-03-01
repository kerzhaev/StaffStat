Attribute VB_Name = "mod_Test_P19"
Option Compare Database

Public Sub InitialSetup_Phase19_Test()
    Dim db As DAO.Database: Set db = CurrentDb
    On Error Resume Next
    ' ??????? ?????? ?????? ?????
    db.Execute "DELETE FROM tbl_Import_Profiles WHERE ProfileID = 1"
    db.Execute "DELETE FROM tbl_Import_Mapping WHERE ProfileID = 1"
    On Error GoTo 0

    ' 1. ??????? ??????? (UID strategy)
    db.Execute "INSERT INTO tbl_Import_Profiles (ProfileID, ProfileName, IdStrategy) VALUES (1, 'Test_Standard', 'UID')"

    ' 2. ??????????? ???????
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tbl_Import_Mapping", dbOpenDynaset)
    AddMap rs, 1, "ID_??????????", "PersonUID"
    AddMap rs, 1, "???_??????", "FullName"
    AddMap rs, 1, "??", "BirthDate"
    rs.Close

    Debug.Print "--- Phase 19 Seed Done ---"
    MsgBox "Seed Completed! Now try importing Excel with headers: 'ID_??????????', '???_??????', '??'", vbInformation
End Sub

Private Sub AddMap(rs As DAO.Recordset, pID As Long, hExc As String, fTarg As String)
    rs.AddNew: rs!ProfileID = pID: rs!ExcelHeader = hExc: rs!targetField = fTarg: rs.Update
End Sub
