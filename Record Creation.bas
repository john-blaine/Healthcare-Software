Attribute VB_Name = "Module3"
Public NoEvents As Boolean

Sub Create_Record_Prep()

    If ThisWorkbook.Sheets(1).Range("D10").Value = "" Then
    MsgBox "Please input all required information."
    Exit Sub
    End If
    
    If ThisWorkbook.Sheets(1).Range("F10").Value = "" Then
    MsgBox "Please input all required information."
    Exit Sub
    End If
    
    Dim DateName As String
    DateName = "Month/Year: " & ThisWorkbook.Sheets("Commands").Range("D10") & " " & ThisWorkbook.Sheets("Commands").Range("F10")
    Dim Sheetname As String
    Sheetname = ThisWorkbook.Sheets(2).Name
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jan" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    MsgBox ("The date entered is in the past. Please enter a valid month and year."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Feb" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    MsgBox ("The date entered is in the past. Please enter a valid month and year."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Mar" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    MsgBox ("The date entered is in the past. Please enter a valid month and year."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If

    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jan" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Feb" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Mar" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Apr" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "May" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jun" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jul" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Aug" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Oct" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Nov" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Dec" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If

    Dim DataObj As New MSForms.DataObject

    'Put a string in the clipboard
    DataObj.SetText ThisWorkbook.Path & "\" & ThisWorkbook.Name & "," & ThisWorkbook.Sheets(1).Range("D10").Value & "," & ThisWorkbook.Sheets(1).Range("F10")
    DataObj.PutInClipboard

    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")

    On Error GoTo Error_Ran_In_VCPI
    
    wsh.Run Chr(34) & "C:\Therapy Grids MatrixCare Interface\Centre Avenue\mc_main_centre_ave.exe" & Chr(34)

    On Error GoTo 0
    
    Exit Sub
    
Error_Ran_In_VCPI:
    MsgBox "An error has occurred. The most likely reason for this is that the therapy grid template is being run inside of VCPI. Please run this therapy grid template outside of VCPI. If this message still appears even when the program is run outside of VCPI, please contact the administrator.", vbCritical, "VCPI Error"
        
End Sub

Sub Import_Worksheet()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim Target_Workbook As Workbook

Dim rng As Range, cell As Range

Dim month_name As String
Dim LMonth As Integer
Dim LYear As Integer
Dim chosenNumber As Integer

Dim last_name As String
Dim first_initial As String

Dim DataObj As New MSForms.DataObject
Dim PatientInfo As String
Dim PatientInfoArray() As String

    DataObj.GetFromClipboard
    PatientInfo = DataObj.GetText()
    
    PatientInfoArray = Split(PatientInfo, ",")
    ThisWorkbook.Sheets(1).Range("D10").Value = PatientInfoArray(1)
    ThisWorkbook.Sheets(1).Range("F10").Value = PatientInfoArray(2)

For Each Target_Workbook In Application.Workbooks
    If Target_Workbook.Name Like "FileServeServlet*.xls" Then
    
        Set Target_Workbook = Target_Workbook

        Exit For
    End If
Next

ThisWorkbook.Sheets(1).Activate

    Dim SheetExists As Boolean

    SheetExists = False
For Each Target_Workbook In Application.Workbooks
    If Target_Workbook.Name Like "FileServeServlet*.xls" Then
            SheetExists = True
            Exit For
        End If
    Next Target_Workbook

If SheetExists = False Then
    Dim answer As Integer
    
    Application.Visible = True
    
    answer = MsgBox("The data import from MatrixCare appears to have failed. Would you like to re-try the import?", vbYesNo + vbQuestion, "Import Prompt")

    If answer = vbYes Then
        Create_Record_Prep
        GoTo End1
    End If
Else
    ThisWorkbook.Sheets(2).Activate
    PatientList.Show
End If

If SheetExists = False Then
    MsgBox ("Since the patient information was not available, please input the patient's name manually.")
    InputName.Show
End If

End1:

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Sub Auto_Populate_Record()

Application.EnableEvents = False

NoEvents = True

Dim AdmitDate As String
Dim MRN As String
Dim MRNArray() As String
Dim Physician As String

Dim Target_Workbook As Workbook

Dim last_name As String
Dim first_initial As String

For Each Target_Workbook In Application.Workbooks
    If Target_Workbook.Name Like "FileServeServlet*.xls" Then

        Set Target_Workbook = Target_Workbook
        
        Exit For
    End If
Next

With Target_Workbook.Sheets(1)

For Each i In .Range("B1:B1000")
    If .Range("B" & i.Row).Value = ThisWorkbook.Sheets(1).Range("A36") And .Range("A" & i.Row).Value = ThisWorkbook.Sheets(1).Range("A23").Value Then
    
        MRN = .Range("E" & i.Row).Value
        MRNArray = Split(MRN, "-")
        MRN = MRNArray(0)
        ThisWorkbook.Sheets(2).Range("A3").Value = "MRN: " & MRN
        
        Physician = .Range("J" & i.Row).Value
        ThisWorkbook.Sheets(2).Range("X3").Value = Physician
        
        If .Range("N" & i.Row).Value = "Acute: Medical Center of the Rockies - Loveland, CO" Or .Range("N" & i.Row).Value = "ACUTE: Medical Center of the Rockies/ OCR Premier - LOVELAND" Or _
        .Range("N" & i.Row).Value = "Acute: Poudre Valley Hospital - Fort Collins, CO" Or .Range("N" & i.Row).Value = "Acute: University of Colorado Hospital - Denver, CO" Or _
        .Range("N" & i.Row).Value = "Acute: Poudre Valley Hospital / OCR Premier - Fort Collins, CO" Or .Range("N" & i.Row).Value = "Rehab:  Rehab Unit at MCR - Loveland, CO" Then
            ThisWorkbook.Sheets(2).CheckBox37.Value = True
        Else
            ThisWorkbook.Sheets(2).CheckBox38.Value = True
        End If
        
        If .Range("F" & i.Row).Value = "MEDICARE A" Then
            ThisWorkbook.Sheets(2).CheckBox1.Value = True
        ElseIf .Range("F" & i.Row).Value = "HUMANA" Then
            ThisWorkbook.Sheets(2).CheckBox2.Value = True
        ElseIf .Range("F" & i.Row).Value = "ANTHEM BCBS" Then
            ThisWorkbook.Sheets(2).CheckBox3.Value = True
        ElseIf .Range("F" & i.Row).Value = "KAISER PERMANENTE" Then
            ThisWorkbook.Sheets(2).CheckBox4.Value = True
        ElseIf .Range("F" & i.Row).Value = "CIGNA" Then
            ThisWorkbook.Sheets(2).CheckBox5.Value = True
            ThisWorkbook.Sheets(2).Range("P2").Value = "Cigna"
        ElseIf .Range("F" & i.Row).Value = "MEDICARE ADVANTAGE" Then
            ThisWorkbook.Sheets(2).CheckBox5.Value = True
            ThisWorkbook.Sheets(2).Range("P2").Value = "Medicare Advantage"
        End If
        
        
    End If
Next i

End With

Target_Workbook.Close

Application.EnableEvents = True

NoEvents = False

Create_Record

End Sub

Sub Create_Record()
Application.DisplayAlerts = False
Application.ScreenUpdating = False

    If ThisWorkbook.Sheets(1).Range("D10").Value = "" Then
    MsgBox "Required information not present. Please contact therapy record administrator."
    Exit Sub
    End If
    
    If ThisWorkbook.Sheets(1).Range("F10").Value = "" Then
    MsgBox "Required information not present. Please contact therapy record administrator."
    Exit Sub
    End If

    With ThisWorkbook.Worksheets(2)
    .Name = ThisWorkbook.Sheets(1).Range("D10") & " " & ThisWorkbook.Sheets(1).Range("F10") & " " & ThisWorkbook.Sheets(1).Range("A20") & ", " & ThisWorkbook.Sheets(1).Range("A19") & " " & "PT"
    End With

Dim DateName As String
DateName = "Month/Year: " & ThisWorkbook.Sheets("Commands").Range("D10") & " " & ThisWorkbook.Sheets("Commands").Range("F10")
Dim Sheetname As String
Sheetname = ThisWorkbook.Sheets(2).Name
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jan" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    MsgBox ("The date entered is in the past. Please enter a valid month and year."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Feb" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    MsgBox ("The date entered is in the past. Please enter a valid month and year."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Mar" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    MsgBox ("The date entered is in the past. Please enter a valid month and year."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If

    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jan" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
    With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Feb" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Mar" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Apr" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "May" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jun" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jul" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Aug" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2016" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Oct" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Nov" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Dec" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    MsgBox ("You have entered a month and year outside of the dates of September 2016-September 2018. Please input dates within that range or contact the administrator."), , "Invalid Dates Entered"
        With ThisWorkbook.Worksheets(2)
    .Name = ("Sheet1")
    End With
    Exit Sub
    End If
    
    ThisWorkbook.Sheets(2).Visible = True
    
    Dim Filepath_Feb2017 As String
    Dim Filepath2_Feb2017 As String
    Dim Filepath3_Feb2017 As String
    Dim Filepath4_Feb2017 As String
    Dim Filepath5_Feb2017 As String
        
    Filepath_Feb2017 = ThisWorkbook.Path & "\Therapy Charts\06. February 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Feb2017 = ThisWorkbook.Path & "\Therapy Charts\06. February 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Feb2017 = ThisWorkbook.Path & "\Therapy Charts\06. February 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Feb2017 = ThisWorkbook.Path & "\Therapy Charts\06. February 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Feb2017 = ThisWorkbook.Path & "\Therapy Charts\06. February 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Feb2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Feb2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Feb2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Feb2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Feb2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Mar2017 As String
    Dim Filepath2_Mar2017 As String
    Dim Filepath3_Mar2017 As String
    Dim Filepath4_Mar2017 As String
    Dim Filepath5_Mar2017 As String
        
    Filepath_Mar2017 = ThisWorkbook.Path & "\Therapy Charts\07. March 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Mar2017 = ThisWorkbook.Path & "\Therapy Charts\07. March 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Mar2017 = ThisWorkbook.Path & "\Therapy Charts\07. March 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Mar2017 = ThisWorkbook.Path & "\Therapy Charts\07. March 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Mar2017 = ThisWorkbook.Path & "\Therapy Charts\07. March 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Mar2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Mar2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Mar2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Mar2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Mar2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Apr2017 As String
    Dim Filepath2_Apr2017 As String
    Dim Filepath3_Apr2017 As String
    Dim Filepath4_Apr2017 As String
    Dim Filepath5_Apr2017 As String
        
    Filepath_Apr2017 = ThisWorkbook.Path & "\Therapy Charts\08. April 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Apr2017 = ThisWorkbook.Path & "\Therapy Charts\08. April 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Apr2017 = ThisWorkbook.Path & "\Therapy Charts\08. April 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Apr2017 = ThisWorkbook.Path & "\Therapy Charts\08. April 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Apr2017 = ThisWorkbook.Path & "\Therapy Charts\08. April 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Apr2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Apr2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Apr2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Apr2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Apr2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_May2017 As String
    Dim Filepath2_May2017 As String
    Dim Filepath3_May2017 As String
    Dim Filepath4_May2017 As String
    Dim Filepath5_May2017 As String
        
    Filepath_May2017 = ThisWorkbook.Path & "\Therapy Charts\09. May 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_May2017 = ThisWorkbook.Path & "\Therapy Charts\09. May 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_May2017 = ThisWorkbook.Path & "\Therapy Charts\09. May 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_May2017 = ThisWorkbook.Path & "\Therapy Charts\09. May 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_May2017 = ThisWorkbook.Path & "\Therapy Charts\09. May 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_May2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_May2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_May2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_May2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_May2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Jun2017 As String
    Dim Filepath2_Jun2017 As String
    Dim Filepath3_Jun2017 As String
    Dim Filepath4_Jun2017 As String
    Dim Filepath5_Jun2017 As String
        
    Filepath_Jun2017 = ThisWorkbook.Path & "\Therapy Charts\10. June 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Jun2017 = ThisWorkbook.Path & "\Therapy Charts\10. June 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Jun2017 = ThisWorkbook.Path & "\Therapy Charts\10. June 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Jun2017 = ThisWorkbook.Path & "\Therapy Charts\10. June 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Jun2017 = ThisWorkbook.Path & "\Therapy Charts\10. June 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Jun2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Jun2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Jun2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Jun2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Jun2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Jul2017 As String
    Dim Filepath2_Jul2017 As String
    Dim Filepath3_Jul2017 As String
    Dim Filepath4_Jul2017 As String
    Dim Filepath5_Jul2017 As String
        
    Filepath_Jul2017 = ThisWorkbook.Path & "\Therapy Charts\11. July 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Jul2017 = ThisWorkbook.Path & "\Therapy Charts\11. July 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Jul2017 = ThisWorkbook.Path & "\Therapy Charts\11. July 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Jul2017 = ThisWorkbook.Path & "\Therapy Charts\11. July 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Jul2017 = ThisWorkbook.Path & "\Therapy Charts\11. July 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Jul2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Jul2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Jul2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Jul2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Jul2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Aug2017 As String
    Dim Filepath2_Aug2017 As String
    Dim Filepath3_Aug2017 As String
    Dim Filepath4_Aug2017 As String
    Dim Filepath5_Aug2017 As String
        
    Filepath_Aug2017 = ThisWorkbook.Path & "\Therapy Charts\12. August 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Aug2017 = ThisWorkbook.Path & "\Therapy Charts\12. August 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Aug2017 = ThisWorkbook.Path & "\Therapy Charts\12. August 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Aug2017 = ThisWorkbook.Path & "\Therapy Charts\12. August 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Aug2017 = ThisWorkbook.Path & "\Therapy Charts\12. August 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Aug2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Aug2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Aug2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Aug2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Aug2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Sep2017 As String
    Dim Filepath2_Sep2017 As String
    Dim Filepath3_Sep2017 As String
    Dim Filepath4_Sep2017 As String
    Dim Filepath5_Sep2017 As String
        
    Filepath_Sep2017 = ThisWorkbook.Path & "\Therapy Charts\13. September 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Sep2017 = ThisWorkbook.Path & "\Therapy Charts\13. September 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Sep2017 = ThisWorkbook.Path & "\Therapy Charts\13. September 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Sep2017 = ThisWorkbook.Path & "\Therapy Charts\13. September 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Sep2017 = ThisWorkbook.Path & "\Therapy Charts\13. September 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Sep2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Sep2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Sep2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Sep2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Sep2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Oct2017 As String
    Dim Filepath2_Oct2017 As String
    Dim Filepath3_Oct2017 As String
    Dim Filepath4_Oct2017 As String
    Dim Filepath5_Oct2017 As String
    
    Filepath_Oct2017 = ThisWorkbook.Path & "\Therapy Charts\14. October 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Oct2017 = ThisWorkbook.Path & "\Therapy Charts\14. October 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Oct2017 = ThisWorkbook.Path & "\Therapy Charts\14. October 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Oct2017 = ThisWorkbook.Path & "\Therapy Charts\14. October 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Oct2017 = ThisWorkbook.Path & "\Therapy Charts\14. October 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Oct2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Oct2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Oct2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Oct2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Oct2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Nov2017 As String
    Dim Filepath2_Nov2017 As String
    Dim Filepath3_Nov2017 As String
    Dim Filepath4_Nov2017 As String
    Dim Filepath5_Nov2017 As String
        
    Filepath_Nov2017 = ThisWorkbook.Path & "\Therapy Charts\15. November 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Nov2017 = ThisWorkbook.Path & "\Therapy Charts\15. November 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Nov2017 = ThisWorkbook.Path & "\Therapy Charts\15. November 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Nov2017 = ThisWorkbook.Path & "\Therapy Charts\15. November 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Nov2017 = ThisWorkbook.Path & "\Therapy Charts\15. November 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Nov2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Nov2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Nov2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Nov2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Nov2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Dec2017 As String
    Dim Filepath2_Dec2017 As String
    Dim Filepath3_Dec2017 As String
    Dim Filepath4_Dec2017 As String
    Dim Filepath5_Dec2017 As String
        
    Filepath_Dec2017 = ThisWorkbook.Path & "\Therapy Charts\16. December 2017\PT\" & Sheetname & ".xlsm"
    Filepath2_Dec2017 = ThisWorkbook.Path & "\Therapy Charts\16. December 2017\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Dec2017 = ThisWorkbook.Path & "\Therapy Charts\16. December 2017\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Dec2017 = ThisWorkbook.Path & "\Therapy Charts\16. December 2017\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Dec2017 = ThisWorkbook.Path & "\Therapy Charts\16. December 2017\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Dec2017)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Dec2017)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Dec2017)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Dec2017)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Dec2017)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Jan2018 As String
    Dim Filepath2_Jan2018 As String
    Dim Filepath3_Jan2018 As String
    Dim Filepath4_Jan2018 As String
    Dim Filepath5_Jan2018 As String
        
    Filepath_Jan2018 = ThisWorkbook.Path & "\Therapy Charts\17. January 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Jan2018 = ThisWorkbook.Path & "\Therapy Charts\17. January 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Jan2018 = ThisWorkbook.Path & "\Therapy Charts\17. January 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Jan2018 = ThisWorkbook.Path & "\Therapy Charts\17. January 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Jan2018 = ThisWorkbook.Path & "\Therapy Charts\17. January 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Jan2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Jan2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Jan2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Jan2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Jan2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If

    Dim Filepath_Feb2018 As String
    Dim Filepath2_Feb2018 As String
    Dim Filepath3_Feb2018 As String
    Dim Filepath4_Feb2018 As String
    Dim Filepath5_Feb2018 As String
        
    Filepath_Feb2018 = ThisWorkbook.Path & "\Therapy Charts\18. February 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Feb2018 = ThisWorkbook.Path & "\Therapy Charts\18. February 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Feb2018 = ThisWorkbook.Path & "\Therapy Charts\18. February 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Feb2018 = ThisWorkbook.Path & "\Therapy Charts\18. February 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Feb2018 = ThisWorkbook.Path & "\Therapy Charts\18. February 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Feb2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Feb2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Feb2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Feb2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Feb2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If

    Dim Filepath_Mar2018 As String
    Dim Filepath2_Mar2018 As String
    Dim Filepath3_Mar2018 As String
    Dim Filepath4_Mar2018 As String
    Dim Filepath5_Mar2018 As String
        
    Filepath_Mar2018 = ThisWorkbook.Path & "\Therapy Charts\19. March 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Mar2018 = ThisWorkbook.Path & "\Therapy Charts\19. March 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Mar2018 = ThisWorkbook.Path & "\Therapy Charts\19. March 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Mar2018 = ThisWorkbook.Path & "\Therapy Charts\19. March 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Mar2018 = ThisWorkbook.Path & "\Therapy Charts\19. March 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Mar2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Mar2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Mar2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Mar2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Mar2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If

    Dim Filepath_Apr2018 As String
    Dim Filepath2_Apr2018 As String
    Dim Filepath3_Apr2018 As String
    Dim Filepath4_Apr2018 As String
    Dim Filepath5_Apr2018 As String
        
    Filepath_Apr2018 = ThisWorkbook.Path & "\Therapy Charts\20. April 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Apr2018 = ThisWorkbook.Path & "\Therapy Charts\20. April 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Apr2018 = ThisWorkbook.Path & "\Therapy Charts\20. April 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Apr2018 = ThisWorkbook.Path & "\Therapy Charts\20. April 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Apr2018 = ThisWorkbook.Path & "\Therapy Charts\20. April 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Apr2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Apr2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Apr2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Apr2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Apr2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_May2018 As String
    Dim Filepath2_May2018 As String
    Dim Filepath3_May2018 As String
    Dim Filepath4_May2018 As String
    Dim Filepath5_May2018 As String
    
    Filepath_May2018 = ThisWorkbook.Path & "\Therapy Charts\21. May 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_May2018 = ThisWorkbook.Path & "\Therapy Charts\21. May 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_May2018 = ThisWorkbook.Path & "\Therapy Charts\21. May 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_May2018 = ThisWorkbook.Path & "\Therapy Charts\21. May 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_May2018 = ThisWorkbook.Path & "\Therapy Charts\21. May 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_May2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_May2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_May2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_May2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_May2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Jun2018 As String
    Dim Filepath2_Jun2018 As String
    Dim Filepath3_Jun2018 As String
    Dim Filepath4_Jun2018 As String
    Dim Filepath5_Jun2018 As String
        
    Filepath_Jun2018 = ThisWorkbook.Path & "\Therapy Charts\22. June 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Jun2018 = ThisWorkbook.Path & "\Therapy Charts\22. June 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Jun2018 = ThisWorkbook.Path & "\Therapy Charts\22. June 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Jun2018 = ThisWorkbook.Path & "\Therapy Charts\22. June 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Jun2018 = ThisWorkbook.Path & "\Therapy Charts\22. June 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Jun2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Jun2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Jun2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Jun2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Jun2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Jul2018 As String
    Dim Filepath2_Jul2018 As String
    Dim Filepath3_Jul2018 As String
    Dim Filepath4_Jul2018 As String
    Dim Filepath5_Jul2018 As String
        
    Filepath_Jul2018 = ThisWorkbook.Path & "\Therapy Charts\23. July 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Jul2018 = ThisWorkbook.Path & "\Therapy Charts\23. July 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Jul2018 = ThisWorkbook.Path & "\Therapy Charts\23. July 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Jul2018 = ThisWorkbook.Path & "\Therapy Charts\23. July 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Jul2018 = ThisWorkbook.Path & "\Therapy Charts\23. July 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Jul2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Jul2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Jul2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Jul2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Jul2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If

    Dim Filepath_Aug2018 As String
    Dim Filepath2_Aug2018 As String
    Dim Filepath3_Aug2018 As String
    Dim Filepath4_Aug2018 As String
    Dim Filepath5_Aug2018 As String
        
    Filepath_Aug2018 = ThisWorkbook.Path & "\Therapy Charts\24. August 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Aug2018 = ThisWorkbook.Path & "\Therapy Charts\24. August 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Aug2018 = ThisWorkbook.Path & "\Therapy Charts\24. August 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Aug2018 = ThisWorkbook.Path & "\Therapy Charts\24. August 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Aug2018 = ThisWorkbook.Path & "\Therapy Charts\24. August 2018\PT\" & Sheetname & " V 5" & ".xlsm"


    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Aug2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Aug2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Aug2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Aug2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Aug2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If
    
    Dim Filepath_Sep2018 As String
    Dim Filepath2_Sep2018 As String
    Dim Filepath3_Sep2018 As String
    Dim Filepath4_Sep2018 As String
    Dim Filepath5_Sep2018 As String
        
    Filepath_Sep2018 = ThisWorkbook.Path & "\Therapy Charts\25. September 2018\PT\" & Sheetname & ".xlsm"
    Filepath2_Sep2018 = ThisWorkbook.Path & "\Therapy Charts\25. September 2018\PT\" & Sheetname & " V 2" & ".xlsm"
    Filepath3_Sep2018 = ThisWorkbook.Path & "\Therapy Charts\25. September 2018\PT\" & Sheetname & " V 3" & ".xlsm"
    Filepath4_Sep2018 = ThisWorkbook.Path & "\Therapy Charts\25. September 2018\PT\" & Sheetname & " V 4" & ".xlsm"
    Filepath5_Sep2018 = ThisWorkbook.Path & "\Therapy Charts\25. September 2018\PT\" & Sheetname & " V 5" & ".xlsm"

    TestStr5 = ""
    On Error Resume Next
    TestStr5 = Dir(Filepath5_Sep2018)
    On Error GoTo 0
    If TestStr5 = "" Then
    Else
    MsgBox ("You are attempting to create a record for a sixth return visit for a patient in the same month. Please contact your Director of Therapy and the Administrator."), , "Re-admission Alert"
    Application.DisplayAlerts = False
    ThisWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If
    
        TestStr4 = ""
        On Error Resume Next
        TestStr4 = Dir(Filepath4_Sep2018)
        On Error GoTo 0
        If TestStr4 = "" Then
        Else
        CreateNewVisit5.Show
        Exit Sub
        End If
        
            TestStr3 = ""
            On Error Resume Next
            TestStr3 = Dir(Filepath3_Sep2018)
            On Error GoTo 0
            If TestStr3 = "" Then
            Else
            CreateNewVisit4.Show
            Exit Sub
            End If
                
                TestStr2 = ""
                On Error Resume Next
                TestStr2 = Dir(Filepath2_Sep2018)
                On Error GoTo 0
                If TestStr2 = "" Then
                Else
                CreateNewVisit3.Show
                Exit Sub
                End If
                
                    TestStr = ""
                    On Error Resume Next
                    TestStr = Dir(Filepath_Sep2018)
                    On Error GoTo 0
                    If TestStr = "" Then
                    Else
                    CreateNewVisit2.Show
                    Exit Sub
                    End If

    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Feb" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\06. February 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Mar" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\07. March 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Apr" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\08. April 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "May" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\09. May 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jun" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\10. June 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jul" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\11. July 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Aug" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\12. August 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Sep" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\13. September 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Oct" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\14. October 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Nov" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\15. November 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Dec" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2017" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\16. December 2017\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jan" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\17. January 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Feb" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\18. February 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Mar" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\19. March 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Apr" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\20. April 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "May" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\21. May 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jun" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\22. June 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Jul" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\23. July 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Aug" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\24. August 2018\PT\" & Sheetname
    GoTo Step2
    End If
    If ThisWorkbook.Sheets("Commands").Range("D10").Value = "Sep" And ThisWorkbook.Sheets("Commands").Range("F10").Value = "2018" Then
    ThisWorkbook.Sheets("Commands").Range("A22").Copy
    ThisWorkbook.Sheets(Sheetname).Range("H3").PasteSpecial
    ThisWorkbook.Sheets("Commands").Range("A21").Copy
    ThisWorkbook.Sheets(Sheetname).Range("O3").PasteSpecial
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\Therapy Charts\25. September 2018\PT\" & Sheetname
    GoTo Step2
    End If
    
Step2:
    ThisWorkbook.Sheets("Commands").Delete

    ThisWorkbook.Save
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    'ThisWorkbook.Close
    
End Sub

