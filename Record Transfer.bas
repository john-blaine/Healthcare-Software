Attribute VB_Name = "Module4"
Sub Roll_Over_Record()

    Dim answer As Integer

    answer = MsgBox("Are you sure you want to roll over this therapy record?", vbYesNo + vbQuestion, "Rollover Prompt")

    If answer = vbYes Then
    Else
    Exit Sub
    End If

Application.DisplayAlerts = False
Application.ScreenUpdating = False
On Error Resume Next
Application.OnTime ThisTime, "Check_Inactivity", Schedule:=False
On Error GoTo 0
ActiveWorkbook.Save

Dim OldSheetName As String
OldSheetName = Sheets(1).Name

    Dim TruePath As String

    ChDrive "G"
    ChDir ThisWorkbook.Path
    
    ChDir ".."
    ChDir ".."
    ChDir ".."
    TruePath = CurDir("G")

If InStr(1, (Sheets(1).Name), "Feb 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Feb ", "Mar ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Mar2017 As String
    Sheetname_Mar2017 = Sheets(1).Name
        
    Filepath_Mar2017 = TruePath & "\Therapy Charts\07. March 2017\PT\" & Sheetname_Mar2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Mar2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the March 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Mar ", "Feb ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Mar 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Mar ", "Apr ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Apr2017 As String
    Sheetname_Apr2017 = Sheets(1).Name
        
    Filepath_Apr2017 = TruePath & "\Therapy Charts\08. April 2017\PT\" & Sheetname_Apr2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Apr2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the April 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Apr ", "Mar ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Apr 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Apr ", "May ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_May2017 As String
    Sheetname_May2017 = Sheets(1).Name
        
    Filepath_May2017 = TruePath & "\Therapy Charts\09. May 2017\PT\" & Sheetname_May2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_May2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the May 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "May ", "Apr ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "May 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "May ", "Jun ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Jun2017 As String
    Sheetname_Jun2017 = Sheets(1).Name
        
    Filepath_Jun2017 = TruePath & "\Therapy Charts\10. June 2017\PT\" & Sheetname_Jun2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Jun2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the June 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jun ", "May ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Jun 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jun ", "Jul ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Jul2017 As String
    Sheetname_Jul2017 = Sheets(1).Name
        
    Filepath_Jul2017 = TruePath & "\Therapy Charts\11. July 2017\PT\" & Sheetname_Jul2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Jul2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the July 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jul ", "Jun ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Jul 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jul ", "Aug ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Aug2017 As String
    Sheetname_Aug2017 = Sheets(1).Name
        
    Filepath_Aug2017 = TruePath & "\Therapy Charts\12. August 2017\PT\" & Sheetname_Aug2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Aug2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the August 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Aug ", "Jul ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Aug 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Aug ", "Sep ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Sep2017 As String
    Sheetname_Sep2017 = Sheets(1).Name
        
    Filepath_Sep2017 = TruePath & "\Therapy Charts\13. September 2017\PT\" & Sheetname_Sep2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Sep2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the September 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Sep ", "Aug ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Sep 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Sep ", "Oct ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Oct2017 As String
    Sheetname_Oct2017 = Sheets(1).Name
        
    Filepath_Oct2017 = TruePath & "\Therapy Charts\14. October 2017\PT\" & Sheetname_Oct2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Oct2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the October 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Oct ", "Sep ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Oct 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Oct ", "Nov ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Nov2017 As String
    Sheetname_Nov2017 = Sheets(1).Name
        
    Filepath_Nov2017 = TruePath & "\Therapy Charts\15. November 2017\PT\" & Sheetname_Nov2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Nov2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the November 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Nov ", "Oct ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Nov 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Nov ", "Dec ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Dec2017 As String
    Sheetname_Dec2017 = Sheets(1).Name
        
    Filepath_Dec2017 = TruePath & "\Therapy Charts\16. December 2017\PT\" & Sheetname_Dec2017 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Dec2017)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the December 2017 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Dec ", "Nov ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Dec 2017") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Dec 2017", "Jan 2018")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Jan2018 As String
    Sheetname_Jan2018 = Sheets(1).Name
        
    Filepath_Jan2018 = TruePath & "\Therapy Charts\17. January 2018\PT\" & Sheetname_Jan2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Jan2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the January 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jan 2018", "Dec 2017")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Jan 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jan ", "Feb ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Feb2018 As String
    Sheetname_Feb2018 = Sheets(1).Name
        
    Filepath_Feb2018 = TruePath & "\Therapy Charts\18. February 2018\PT\" & Sheetname_Feb2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Feb2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the February 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Feb ", "Jan ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Feb 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Feb ", "Mar ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Mar2018 As String
    Sheetname_Mar2018 = Sheets(1).Name
        
    Filepath_Mar2018 = TruePath & "\Therapy Charts\19. March 2018\PT\" & Sheetname_Mar2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Mar2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the March 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Mar ", "Feb ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Mar 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Mar ", "Apr ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Apr2018 As String
    Sheetname_Apr2018 = Sheets(1).Name
        
    Filepath_Apr2018 = TruePath & "\Therapy Charts\20. April 2018\PT\" & Sheetname_Apr2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Apr2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the April 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Apr ", "Mar ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Apr 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Apr ", "May ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_May2018 As String
    Sheetname_May2018 = Sheets(1).Name
        
    Filepath_May2018 = TruePath & "\Therapy Charts\21. May 2018\PT\" & Sheetname_May2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_May2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the May 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "May ", "Apr ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "May 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "May ", "Jun ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Jun2018 As String
    Sheetname_Jun2018 = Sheets(1).Name
        
    Filepath_Jun2018 = TruePath & "\Therapy Charts\22. June 2018\PT\" & Sheetname_Jun2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Jun2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the June 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jun ", "May ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Jun 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jun ", "Jul ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Jul2018 As String
    Sheetname_Jul2018 = Sheets(1).Name
        
    Filepath_Jul2018 = TruePath & "\Therapy Charts\23. July 2018\PT\" & Sheetname_Jul2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Jul2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the July 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jul ", "Jun ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Jul 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Jul ", "Aug ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Aug2018 As String
    Sheetname_Aug2018 = Sheets(1).Name
        
    Filepath_Aug2018 = TruePath & "\Therapy Charts\24. August 2018\PT\" & Sheetname_Aug2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Aug2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the August 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Aug ", "Jul ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If

With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Aug 2018") > 0 Then
ActiveWorkbook.Save
Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Aug ", "Sep ")
Sheets(1).Name = Replace(Sheets(1).Name, " V 2", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 3", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 4", "")
Sheets(1).Name = Replace(Sheets(1).Name, " V 5", "")

    Dim Sheetname_Sep2018 As String
    Sheetname_Sep2018 = Sheets(1).Name
        
    Filepath_Sep2018 = TruePath & "\Therapy Charts\25. September 2018\PT\" & Sheetname_Sep2018 & ".xlsm"

    TestStr = ""
    On Error Resume Next
    TestStr = Dir(Filepath_Sep2018)
    On Error GoTo 0
    If TestStr = "" Then
    Else
    MsgBox ("A record for this patient and this date has already been created. Please check the September 2018 folder."), , "Record Already Exists"
    Sheets(1).Name = WorksheetFunction.Substitute(Sheets(1).Name, "Sep ", "Aug ")
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Exit Sub
    End If


With ActiveWorkbook.Worksheets(1)

End With
GoTo Step_2
End If

If InStr(1, (Sheets(1).Name), "Sep 2018") > 0 Then
MsgBox ("The record cannot be rolled over. Further Excel programming must be completed. Please contact the administrator."), , "Requested Operation Outside Of Programmed Dates"
Exit Sub
End If

Step_2:
Dim WorkbookName As String
WorkbookName = ActiveWorkbook.Name

Dim Sheetname As String
Sheetname = Sheets(1).Name

If InStr(1, (Sheets(1).Name), "Feb 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\06. February 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\05. January 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\05. January 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\05. January 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True

End If

If InStr(1, (Sheets(1).Name), "Mar 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\07. March 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\06. February 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\06. February 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\06. February 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True

End If

If InStr(1, (Sheets(1).Name), "Apr 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\08. April 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\07. March 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\07. March 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\07. March 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "May 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\09. May 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\08. April 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\08. April 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\08. April 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True

End If

If InStr(1, (Sheets(1).Name), "Jun 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\10. June 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\09. May 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\09. May 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\09. May 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True

End If

If InStr(1, (Sheets(1).Name), "Jul 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\11. July 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\10. June 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\10. June 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\10. June 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True

End If

If InStr(1, (Sheets(1).Name), "Aug 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\12. August 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\11. July 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\11. July 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\11. July 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Sep 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\13. September 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\12. August 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\12. August 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\12. August 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True

End If

If InStr(1, (Sheets(1).Name), "Oct 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\14. October 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\13. September 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\13. September 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\13. September 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Nov 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\15. November 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\14. October 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\14. October 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\14. October 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Dec 2017") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\16. December 2017\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\15. November 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\15. November 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\15. November 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Jan 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\17. January 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\16. December 2017\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\16. December 2017\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\16. December 2017\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Feb 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\18. February 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\17. January 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\17. January 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\17. January 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Mar 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\19. March 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\18. February 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\18. February 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\18. February 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Apr 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\20. April 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\19. March 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\19. March 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\19. March 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "May 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\21. May 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\20. April 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\20. April 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\20. April 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Jun 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\22. June 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\21. May 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\21. May 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\21. May 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Jul 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\23. July 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\22. June 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\22. June 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\22. June 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Aug 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\24. August 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\23. July 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\23. July 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\23. July 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

If InStr(1, (Sheets(1).Name), "Sep 2018") > 0 Then
ActiveWorkbook.SaveAs TruePath & "\Therapy Charts\25. September 2018\PT\" & Sheetname
Sheets(1).Unprotect "healingarts"
Workbooks.Open (TruePath & "\Therapy Charts\24. August 2018\PT\" & OldSheetName & ".xlsm")
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Range("PreviousMonthVisits").Formula = "='" & TruePath & "\Therapy Charts\24. August 2018\PT\" & OldSheetName & ".xlsm" & "'!CurrentMonthVisits" & " " & "+" & " " & "'" & TruePath & "\Therapy Charts\24. August 2018\PT\" & OldSheetName & ".xlsm" & "'!PreviousMonthVisits"
Workbooks(OldSheetName & ".xlsm").Activate
ActiveWorkbook.Close
Workbooks(Sheetname & ".xlsm").Activate
Sheets(1).Protect "healingarts", AllowFormattingCells:=True
End If

With ActiveWorkbook.Worksheets(1)
On Error Resume Next
Worksheets(1).Range("ColumnStartB:Column2EndAF").ClearContents
ActiveSheet.Unprotect "healingarts"
ActiveWorkbook.Sheets(1).Range("InitialsB1:InitialsAF3").Locked = False
Worksheets(1).Range("Column2StartB:ColumnEndAF").ClearContents
Worksheets(1).Range("InitialsB1:InitialsAF3").ClearContents
Reset_Status_Codes
ActiveWorkbook.Sheets(1).Range("InitialsB1:InitialsAF3").Locked = True

For Each m_cell In Range("MergeAreaTitles")
    m_cell.MergeArea.ClearContents
    Next m_cell
    
ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
Application.OnTime ThisTime, "Check_Inactivity", Schedule:=False
End With

ActiveWorkbook.Save

MsgBox "Record Rollover Successfully Completed"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Private Sub Reset_Status_Codes()

On Error Resume Next

Range("A9:A27").Interior.ColorIndex = 0

Worksheets(1).Range("InitialsB1").Comment.Delete
Worksheets(1).Range("InitialsB2").Comment.Delete
Worksheets(1).Range("InitialsB3").Comment.Delete
Worksheets(1).Range("InitialsC1").Comment.Delete
Worksheets(1).Range("InitialsC2").Comment.Delete
Worksheets(1).Range("InitialsC3").Comment.Delete
Worksheets(1).Range("InitialsD1").Comment.Delete
Worksheets(1).Range("InitialsD2").Comment.Delete
Worksheets(1).Range("InitialsD3").Comment.Delete
Worksheets(1).Range("InitialsE1").Comment.Delete
Worksheets(1).Range("InitialsE2").Comment.Delete
Worksheets(1).Range("InitialsE3").Comment.Delete
Worksheets(1).Range("InitialsF1").Comment.Delete
Worksheets(1).Range("InitialsF2").Comment.Delete
Worksheets(1).Range("InitialsF3").Comment.Delete
Worksheets(1).Range("InitialsG1").Comment.Delete
Worksheets(1).Range("InitialsG2").Comment.Delete
Worksheets(1).Range("InitialsG3").Comment.Delete
Worksheets(1).Range("InitialsH1").Comment.Delete
Worksheets(1).Range("InitialsH2").Comment.Delete
Worksheets(1).Range("InitialsH3").Comment.Delete
Worksheets(1).Range("InitialsI1").Comment.Delete
Worksheets(1).Range("InitialsI2").Comment.Delete
Worksheets(1).Range("InitialsI3").Comment.Delete
Worksheets(1).Range("InitialsJ1").Comment.Delete
Worksheets(1).Range("InitialsJ2").Comment.Delete
Worksheets(1).Range("InitialsJ3").Comment.Delete
Worksheets(1).Range("InitialsK1").Comment.Delete
Worksheets(1).Range("InitialsK2").Comment.Delete
Worksheets(1).Range("InitialsK3").Comment.Delete
Worksheets(1).Range("InitialsL1").Comment.Delete
Worksheets(1).Range("InitialsL2").Comment.Delete
Worksheets(1).Range("InitialsL3").Comment.Delete
Worksheets(1).Range("InitialsM1").Comment.Delete
Worksheets(1).Range("InitialsM2").Comment.Delete
Worksheets(1).Range("InitialsM3").Comment.Delete
Worksheets(1).Range("InitialsN1").Comment.Delete
Worksheets(1).Range("InitialsN2").Comment.Delete
Worksheets(1).Range("InitialsN3").Comment.Delete
Worksheets(1).Range("InitialsO1").Comment.Delete
Worksheets(1).Range("InitialsO2").Comment.Delete
Worksheets(1).Range("InitialsO3").Comment.Delete
Worksheets(1).Range("InitialsP1").Comment.Delete
Worksheets(1).Range("InitialsP2").Comment.Delete
Worksheets(1).Range("InitialsP3").Comment.Delete
Worksheets(1).Range("InitialsQ1").Comment.Delete
Worksheets(1).Range("InitialsQ2").Comment.Delete
Worksheets(1).Range("InitialsQ3").Comment.Delete
Worksheets(1).Range("InitialsR1").Comment.Delete
Worksheets(1).Range("InitialsR2").Comment.Delete
Worksheets(1).Range("InitialsR3").Comment.Delete
Worksheets(1).Range("InitialsS1").Comment.Delete
Worksheets(1).Range("InitialsS2").Comment.Delete
Worksheets(1).Range("InitialsS3").Comment.Delete
Worksheets(1).Range("InitialsT1").Comment.Delete
Worksheets(1).Range("InitialsT2").Comment.Delete
Worksheets(1).Range("InitialsT3").Comment.Delete
Worksheets(1).Range("InitialsU1").Comment.Delete
Worksheets(1).Range("InitialsU2").Comment.Delete
Worksheets(1).Range("InitialsU3").Comment.Delete
Worksheets(1).Range("InitialsV1").Comment.Delete
Worksheets(1).Range("InitialsV2").Comment.Delete
Worksheets(1).Range("InitialsV3").Comment.Delete
Worksheets(1).Range("InitialsW1").Comment.Delete
Worksheets(1).Range("InitialsW2").Comment.Delete
Worksheets(1).Range("InitialsW3").Comment.Delete
Worksheets(1).Range("InitialsX1").Comment.Delete
Worksheets(1).Range("InitialsX2").Comment.Delete
Worksheets(1).Range("InitialsX3").Comment.Delete
Worksheets(1).Range("InitialsY1").Comment.Delete
Worksheets(1).Range("InitialsY2").Comment.Delete
Worksheets(1).Range("InitialsY3").Comment.Delete
Worksheets(1).Range("InitialsZ1").Comment.Delete
Worksheets(1).Range("InitialsZ2").Comment.Delete
Worksheets(1).Range("InitialsZ3").Comment.Delete
Worksheets(1).Range("InitialsAA1").Comment.Delete
Worksheets(1).Range("InitialsAA2").Comment.Delete
Worksheets(1).Range("InitialsAA3").Comment.Delete
Worksheets(1).Range("InitialsAB1").Comment.Delete
Worksheets(1).Range("InitialsAB2").Comment.Delete
Worksheets(1).Range("InitialsAB3").Comment.Delete
Worksheets(1).Range("InitialsAC1").Comment.Delete
Worksheets(1).Range("InitialsAC2").Comment.Delete
Worksheets(1).Range("InitialsAC3").Comment.Delete
Worksheets(1).Range("InitialsAD1").Comment.Delete
Worksheets(1).Range("InitialsAD2").Comment.Delete
Worksheets(1).Range("InitialsAD3").Comment.Delete
Worksheets(1).Range("InitialsAE1").Comment.Delete
Worksheets(1).Range("InitialsAE2").Comment.Delete
Worksheets(1).Range("InitialsAE3").Comment.Delete
Worksheets(1).Range("InitialsAF1").Comment.Delete
Worksheets(1).Range("InitialsAF2").Comment.Delete
Worksheets(1).Range("InitialsAF3").Comment.Delete

Range("TreatmentMinutesB").Value = "=sum(Column2StartB,Column3EndB, Column3StartB:ColumnEndB)"
Range("TreatmentMinutesC").Value = "=sum(Column2StartC,Column3EndC, Column3StartC:ColumnEndC)"
Range("TreatmentMinutesD").Value = "=sum(Column2StartD,Column3EndD, Column3StartD:ColumnEndD)"
Range("TreatmentMinutesE").Value = "=sum(Column2StartE,Column3EndE, Column3StartE:ColumnEndE)"
Range("TreatmentMinutesF").Value = "=sum(Column2StartF,Column3EndF, Column3StartF:ColumnEndF)"
Range("TreatmentMinutesG").Value = "=sum(Column2StartG,Column3EndG, Column3StartG:ColumnEndG)"
Range("TreatmentMinutesH").Value = "=sum(Column2StartH,Column3EndH, Column3StartH:ColumnEndH)"
Range("TreatmentMinutesI").Value = "=sum(Column2StartI,Column3EndI, Column3StartI:ColumnEndI)"
Range("TreatmentMinutesJ").Value = "=sum(Column2StartJ,Column3EndJ, Column3StartJ:ColumnEndJ)"
Range("TreatmentMinutesK").Value = "=sum(Column2StartK,Column3EndK, Column3StartK:ColumnEndK)"
Range("TreatmentMinutesL").Value = "=sum(Column2StartL,Column3EndL, Column3StartL:ColumnEndL)"
Range("TreatmentMinutesM").Value = "=sum(Column2StartM,Column3EndM, Column3StartM:ColumnEndM)"
Range("TreatmentMinutesN").Value = "=sum(Column2StartN,Column3EndN, Column3StartN:ColumnEndN)"
Range("TreatmentMinutesO").Value = "=sum(Column2StartO,Column3EndO, Column3StartO:ColumnEndO)"
Range("TreatmentMinutesP").Value = "=sum(Column2StartP,Column3EndP, Column3StartP:ColumnEndP)"
Range("TreatmentMinutesQ").Value = "=sum(Column2StartQ,Column3EndQ, Column3StartQ:ColumnEndQ)"
Range("TreatmentMinutesR").Value = "=sum(Column2StartR,Column3EndR, Column3StartR:ColumnEndR)"
Range("TreatmentMinutesS").Value = "=sum(Column2StartS,Column3EndS, Column3StartS:ColumnEndS)"
Range("TreatmentMinutesT").Value = "=sum(Column2StartT,Column3EndT, Column3StartT:ColumnEndT)"
Range("TreatmentMinutesU").Value = "=sum(Column2StartU,Column3EndU, Column3StartU:ColumnEndU)"
Range("TreatmentMinutesV").Value = "=sum(Column2StartV,Column3EndV, Column3StartV:ColumnEndV)"
Range("TreatmentMinutesW").Value = "=sum(Column2StartW,Column3EndW, Column3StartW:ColumnEndW)"
Range("TreatmentMinutesX").Value = "=sum(Column2StartX,Column3EndX, Column3StartX:ColumnEndX)"
Range("TreatmentMinutesY").Value = "=sum(Column2StartY,Column3EndY, Column3StartY:ColumnEndY)"
Range("TreatmentMinutesZ").Value = "=sum(Column2StartZ,Column3EndZ, Column3StartZ:ColumnEndZ)"
Range("TreatmentMinutesAA").Value = "=sum(Column2StartAA,Column3EndAA, Column3StartAA:ColumnEndAA)"
Range("TreatmentMinutesAB").Value = "=sum(Column2StartAB,Column3EndAB, Column3StartAB:ColumnEndAB)"
Range("TreatmentMinutesAC").Value = "=sum(Column2StartAC,Column3EndAC, Column3StartAC:ColumnEndAC)"
Range("TreatmentMinutesAD").Value = "=sum(Column2StartAD,Column3EndAD, Column3StartAD:ColumnEndAD)"
Range("TreatmentMinutesAE").Value = "=sum(Column2StartAE,Column3EndAE, Column3StartAE:ColumnEndAE)"
Range("TreatmentMinutesAF").Value = "=sum(Column2StartAF,Column3EndAF, Column3StartAF:ColumnEndAF)"

Range("TotalTreatmentMinutesB").Value = "=SUM(B28, B17, B15)"
Range("TotalTreatmentMinutesC").Value = "=SUM(C28, C17, C15)"
Range("TotalTreatmentMinutesD").Value = "=SUM(D28, D17, D15)"
Range("TotalTreatmentMinutesE").Value = "=SUM(E28, E17, E15)"
Range("TotalTreatmentMinutesF").Value = "=SUM(F28, F17, F15)"
Range("TotalTreatmentMinutesG").Value = "=SUM(G28, G17, G15)"
Range("TotalTreatmentMinutesH").Value = "=SUM(H28, H17, H15)"
Range("TotalTreatmentMinutesI").Value = "=SUM(I28, I17, I15)"
Range("TotalTreatmentMinutesJ").Value = "=SUM(J28, J17, J15)"
Range("TotalTreatmentMinutesK").Value = "=SUM(K28, K17, K15)"
Range("TotalTreatmentMinutesL").Value = "=SUM(L28, L17, L15)"
Range("TotalTreatmentMinutesM").Value = "=SUM(M28, M17, M15)"
Range("TotalTreatmentMinutesN").Value = "=SUM(N28, N17, N15)"
Range("TotalTreatmentMinutesO").Value = "=SUM(O28, O17, O15)"
Range("TotalTreatmentMinutesP").Value = "=SUM(P28, P17, P15)"
Range("TotalTreatmentMinutesQ").Value = "=SUM(Q28, Q17, Q15)"
Range("TotalTreatmentMinutesR").Value = "=SUM(R28, R17, R15)"
Range("TotalTreatmentMinutesS").Value = "=SUM(S28, S17, S15)"
Range("TotalTreatmentMinutesT").Value = "=SUM(T28, T17, T15)"
Range("TotalTreatmentMinutesU").Value = "=SUM(U28, U17, U15)"
Range("TotalTreatmentMinutesV").Value = "=SUM(V28, V17, V15)"
Range("TotalTreatmentMinutesW").Value = "=SUM(W28, W17, W15)"
Range("TotalTreatmentMinutesX").Value = "=SUM(X28, X17, X15)"
Range("TotalTreatmentMinutesY").Value = "=SUM(Y28, Y17, Y15)"
Range("TotalTreatmentMinutesZ").Value = "=SUM(Z28, Z17, Z15)"
Range("TotalTreatmentMinutesAA").Value = "=SUM(AA28, AA17, AA15)"
Range("TotalTreatmentMinutesAB").Value = "=SUM(AB28, AB17, AB15)"
Range("TotalTreatmentMinutesAC").Value = "=SUM(AC28, AC17, AC15)"
Range("TotalTreatmentMinutesAD").Value = "=SUM(AD28, AD17, AD15)"
Range("TotalTreatmentMinutesAE").Value = "=SUM(AE28, AE17, AE15)"
Range("TotalTreatmentMinutesAF").Value = "=SUM(AF28, AF17, AF15)"

ActiveSheet.Shapes("Discharge1").Visible = False
ActiveSheet.Shapes("Discharge2").Visible = False
ActiveSheet.Shapes("Discharge3").Visible = False
ActiveSheet.Shapes("Discharge4").Visible = False
ActiveSheet.Shapes("Discharge5").Visible = False
ActiveSheet.Shapes("Discharge6").Visible = False
ActiveSheet.Shapes("Discharge7").Visible = False
ActiveSheet.Shapes("Discharge8").Visible = False
ActiveSheet.Shapes("Discharge9").Visible = False
ActiveSheet.Shapes("Discharge10").Visible = False
ActiveSheet.Shapes("Discharge11").Visible = False
ActiveSheet.Shapes("Discharge12").Visible = False
ActiveSheet.Shapes("Discharge13").Visible = False
ActiveSheet.Shapes("Discharge14").Visible = False
ActiveSheet.Shapes("Discharge15").Visible = False
ActiveSheet.Shapes("Discharge16").Visible = False
ActiveSheet.Shapes("Discharge17").Visible = False
ActiveSheet.Shapes("Discharge18").Visible = False
ActiveSheet.Shapes("Discharge19").Visible = False
ActiveSheet.Shapes("Discharge20").Visible = False
ActiveSheet.Shapes("Discharge21").Visible = False
ActiveSheet.Shapes("Discharge22").Visible = False
ActiveSheet.Shapes("Discharge23").Visible = False
ActiveSheet.Shapes("Discharge24").Visible = False
ActiveSheet.Shapes("Discharge25").Visible = False
ActiveSheet.Shapes("Discharge26").Visible = False
ActiveSheet.Shapes("Discharge27").Visible = False
ActiveSheet.Shapes("Discharge28").Visible = False
ActiveSheet.Shapes("Discharge29").Visible = False
ActiveSheet.Shapes("Discharge30").Visible = False

ActiveSheet.CheckBox6.Value = False
ActiveSheet.CheckBox7.Value = False
ActiveSheet.CheckBox8.Value = False
ActiveSheet.CheckBox9.Value = False
ActiveSheet.CheckBox10.Value = False
ActiveSheet.CheckBox11.Value = False
ActiveSheet.CheckBox12.Value = False
ActiveSheet.CheckBox13.Value = False
ActiveSheet.CheckBox14.Value = False
ActiveSheet.CheckBox15.Value = False
ActiveSheet.CheckBox16.Value = False
ActiveSheet.CheckBox17.Value = False
ActiveSheet.CheckBox18.Value = False
ActiveSheet.CheckBox19.Value = False
ActiveSheet.CheckBox20.Value = False
ActiveSheet.CheckBox21.Value = False
ActiveSheet.CheckBox22.Value = False
ActiveSheet.CheckBox23.Value = False
ActiveSheet.CheckBox24.Value = False
ActiveSheet.CheckBox25.Value = False
ActiveSheet.CheckBox26.Value = False
ActiveSheet.CheckBox27.Value = False
ActiveSheet.CheckBox28.Value = False
ActiveSheet.CheckBox29.Value = False
ActiveSheet.CheckBox30.Value = False
ActiveSheet.CheckBox31.Value = False
ActiveSheet.CheckBox32.Value = False
ActiveSheet.CheckBox33.Value = False
ActiveSheet.CheckBox34.Value = False
ActiveSheet.CheckBox35.Value = False
ActiveSheet.CheckBox36.Value = False

Worksheets(1).Range("TreatmentFrequency1:TreatmentFrequency31").ClearContents

End Sub
