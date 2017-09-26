Attribute VB_Name = "Module5"
Public Initials As Variant
Public SelectedCell As Variant
Public Day_Date As Integer

Sub CheckAll()

If ActiveSheet.Shapes("Discharge1").Visible = True Or ActiveSheet.Shapes("Discharge25").Visible = True Or _
    ActiveSheet.Shapes("Discharge2").Visible = True Or ActiveSheet.Shapes("Discharge26").Visible = True Or _
    ActiveSheet.Shapes("Discharge3").Visible = True Or ActiveSheet.Shapes("Discharge27").Visible = True Or _
    ActiveSheet.Shapes("Discharge4").Visible = True Or ActiveSheet.Shapes("Discharge28").Visible = True Or _
    ActiveSheet.Shapes("Discharge5").Visible = True Or ActiveSheet.Shapes("Discharge29").Visible = True Or _
    ActiveSheet.Shapes("Discharge6").Visible = True Or ActiveSheet.Shapes("Discharge30").Visible = True Or _
    ActiveSheet.Shapes("Discharge7").Visible = True Or _
    ActiveSheet.Shapes("Discharge8").Visible = True Or _
    ActiveSheet.Shapes("Discharge9").Visible = True Or _
    ActiveSheet.Shapes("Discharge10").Visible = True Or _
    ActiveSheet.Shapes("Discharge11").Visible = True Or _
    ActiveSheet.Shapes("Discharge12").Visible = True Or _
    ActiveSheet.Shapes("Discharge13").Visible = True Or _
    ActiveSheet.Shapes("Discharge14").Visible = True Or _
    ActiveSheet.Shapes("Discharge15").Visible = True Or _
    ActiveSheet.Shapes("Discharge16").Visible = True Or _
    ActiveSheet.Shapes("Discharge17").Visible = True Or _
    ActiveSheet.Shapes("Discharge18").Visible = True Or _
    ActiveSheet.Shapes("Discharge19").Visible = True Or _
    ActiveSheet.Shapes("Discharge20").Visible = True Or _
    ActiveSheet.Shapes("Discharge21").Visible = True Or _
    ActiveSheet.Shapes("Discharge22").Visible = True Or _
    ActiveSheet.Shapes("Discharge23").Visible = True Or _
    ActiveSheet.Shapes("Discharge24").Visible = True Then
    
    MsgBox "This patient is indicated as having been discharged. If this is an error, please press the D/C button again and re-enter your initials."
    Initials = ""
    Exit Sub
    End If

    Initials = InputBox("Please enter your initials. By doing so, you attest that you are the clinician providing care on this date to the patient on record.")
    Initials = Replace(Initials, " ", "")
    ThisWorkbook.Sheets(1).Range("Initials1").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials1").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials2").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials2").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials3").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials3").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials4").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials4").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials5").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials5").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials6").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials6").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials7").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials7").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials8").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials8").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials9").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials9").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials10").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials10").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials11").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials11").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials12").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials12").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials13").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials13").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials14").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials14").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials15").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials15").Value), " ", "")
    ThisWorkbook.Sheets(1).Range("Initials16").Value = Replace((ThisWorkbook.Sheets(1).Range("Initials16").Value), " ", "")
    If Initials = "" Then
        Exit Sub
        End If
        
            If Initials = (ThisWorkbook.Sheets(1).Range("Initials1").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials2").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials3").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials4").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials5").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials6").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials7").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials8").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials9").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials10").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials11").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials12").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials13").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials14").Value) Or Initials = (ThisWorkbook.Sheets(1).Range("Initials15").Value) _
            Or Initials = (ThisWorkbook.Sheets(1).Range("Initials16").Value) Then
            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials1")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName1")) = True Then
            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
            Initials = ""
            Exit Sub
            End If
                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials1")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title1")) = True Then
                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                Initials = ""
                Exit Sub
                End If
                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials2")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName2")) = True Then
                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                    Initials = ""
                    Exit Sub
                    End If
                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials2")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title2")) = True Then
                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                        Initials = ""
                        Exit Sub
                        End If
                            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials3")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName3")) = True Then
                            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                            Initials = ""
                            Exit Sub
                            End If
                                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials3")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title3")) = True Then
                                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                Initials = ""
                                Exit Sub
                                End If
                                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials4")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName4")) = True Then
                                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                    Initials = ""
                                    Exit Sub
                                    End If
                                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials4")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title4")) = True Then
                                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                        Initials = ""
                                        Exit Sub
                                        End If
                                            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials5")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName5")) = True Then
                                            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                            Initials = ""
                                            Exit Sub
                                            End If
                                                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials5")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title5")) = True Then
                                                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                Initials = ""
                                                Exit Sub
                                                End If
                                                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials6")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName6")) = True Then
                                                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                    Initials = ""
                                                    Exit Sub
                                                    End If
                                                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials6")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title6")) = True Then
                                                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                        Initials = ""
                                                        Exit Sub
                                                        End If
                                                            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials7")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName7")) = True Then
                                                            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                            Initials = ""
                                                            Exit Sub
                                                            End If
                                                                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials7")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title7")) = True Then
                                                                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                Initials = ""
                                                                Exit Sub
                                                                End If
                                                                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials8")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName8")) = True Then
                                                                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                    Initials = ""
                                                                    Exit Sub
                                                                    End If
                                                                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials8")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title8")) = True Then
                                                                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                        Initials = ""
                                                                        Exit Sub
                                                                        End If
                                                                            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials9")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName9")) = True Then
                                                                            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                            Initials = ""
                                                                            Exit Sub
                                                                            End If
                                                                                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials9")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title9")) = True Then
                                                                                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                Initials = ""
                                                                                Exit Sub
                                                                                End If
                                                                                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials10")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName10")) = True Then
                                                                                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                    Initials = ""
                                                                                    Exit Sub
                                                                                    End If
                                                                                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials10")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title10")) = True Then
                                                                                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                        Initials = ""
                                                                                        Exit Sub
                                                                                        End If
                                                                                            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials11")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName11")) = True Then
                                                                                            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                            Initials = ""
                                                                                            Exit Sub
                                                                                            End If
                                                                                                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials11")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title11")) = True Then
                                                                                                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                Initials = ""
                                                                                                Exit Sub
                                                                                                End If
                                                                                                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials12")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName12")) = True Then
                                                                                                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                    Initials = ""
                                                                                                    Exit Sub
                                                                                                    End If
                                                                                                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials12")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title12")) = True Then
                                                                                                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                        Initials = ""
                                                                                                        Exit Sub
                                                                                                        End If
                                                                                                            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials13")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName13")) = True Then
                                                                                                            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                            Initials = ""
                                                                                                            Exit Sub
                                                                                                            End If
                                                                                                                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials13")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title13")) = True Then
                                                                                                                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                                Initials = ""
                                                                                                                Exit Sub
                                                                                                                End If
                                                                                                                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials14")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName14")) = True Then
                                                                                                                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                                    Initials = ""
                                                                                                                    Exit Sub
                                                                                                                    End If
                                                                                                                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials14")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title14")) = True Then
                                                                                                                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                                        Initials = ""
                                                                                                                        Exit Sub
                                                                                                                        End If
                                                                                                                            If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials15")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("TherapistName15")) = True Then
                                                                                                                            MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                                            Initials = ""
                                                                                                                            Exit Sub
                                                                                                                            End If
                                                                                                                                If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials15")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title15")) = True Then
                                                                                                                                MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                                                Initials = ""
                                                                                                                                Exit Sub
                                                                                                                                End If
                                                                                                                                    If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials16")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title16")) = True Then
                                                                                                                                    MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                                                    Initials = ""
                                                                                                                                    Exit Sub
                                                                                                                                    End If
                                                                                                                                        If IsEmpty(ThisWorkbook.Sheets(1).Range("Initials16")) = False And IsEmpty(ThisWorkbook.Sheets(1).Range("Title16")) = True Then
                                                                                                                                        MsgBox "A name and/or title has not been entered in the box below - this must be completed before entering initials."
                                                                                                                                        Initials = ""
                                                                                                                                        Exit Sub
                                                                                                                                        End If
            Else
            MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
            Initials = ""
            Exit Sub
            End If

Dim Year_Date As String
Dim Month_Date As String
Dim Test_Date_2

If InStr(1, (Sheets(1).Name), "Jan ") > 0 Then
Month_Date = "1"
End If

If InStr(1, (Sheets(1).Name), "Feb ") > 0 Then
Month_Date = "2"
End If

If InStr(1, (Sheets(1).Name), "Mar ") > 0 Then
Month_Date = "3"
End If

If InStr(1, (Sheets(1).Name), "Apr ") > 0 Then
Month_Date = "4"
End If

If InStr(1, (Sheets(1).Name), "May ") > 0 Then
Month_Date = "5"
End If

If InStr(1, (Sheets(1).Name), "Jun ") > 0 Then
Month_Date = "6"
End If

If InStr(1, (Sheets(1).Name), "Jul ") > 0 Then
Month_Date = "7"
End If

If InStr(1, (Sheets(1).Name), "Aug ") > 0 Then
Month_Date = "8"
End If

If InStr(1, (Sheets(1).Name), "Sep ") > 0 Then
Month_Date = "9"
End If

If InStr(1, (Sheets(1).Name), "Oct ") > 0 Then
Month_Date = "10"
End If

If InStr(1, (Sheets(1).Name), "Nov ") > 0 Then
Month_Date = "11"
End If

If InStr(1, (Sheets(1).Name), "Dec ") > 0 Then
Month_Date = "12"
End If

If InStr(1, (Sheets(1).Name), "2017") > 0 Then
Year_Date = "2017"
End If

If InStr(1, (Sheets(1).Name), "2018") > 0 Then
Year_Date = "2018"
End If

If InStr(1, (Sheets(1).Name), "2019") > 0 Then
Year_Date = "2019"
End If

If Year(Date) < Year_Date Then
MsgBox "You are attempting to chart on a day that is in the future. Charting ahead of time is currently prohibited. Please chart on today's date or a date further in the past.", , "Charting Ahead Not Allowed"
Initials = ""
Exit Sub
End If

If Year(Date) = Year_Date Then
If Month(Date) < Month_Date Then
MsgBox "You are attempting to chart on a day that is in the future. Charting ahead of time is currently prohibited. Please chart on today's date or a date further in the past.", , "Charting Ahead Not Allowed"
Initials = ""
Exit Sub
End If
End If

If Year(Date) = Year_Date Then
If Month(Date) = Month_Date Then
If Day(Date) < Day_Date Then
MsgBox "You are attempting to chart on a day that is in the future. Charting ahead of time is currently prohibited. Please chart on today's date or a date further in the past.", , "Charting Ahead Not Allowed"
Initials = ""
Exit Sub
End If
End If
End If

End Sub

Sub Protect_InitialsB1()

Day_Date = "1"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If
    
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsB1").Locked = False
    SelectedCell = "B"
    
        CheckAll
    
    Range("InitialsB1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsB1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsB1").Locked = False
    SelectedCell = "B"
    
        CheckAll
    
    Range("InitialsB1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsB1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub
    
Sub Protect_InitialsB2()

Day_Date = "1"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsB2").Locked = False
    SelectedCell = "B"
    
        CheckAll
    
    Range("InitialsB2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsB2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsB2").Locked = False
    SelectedCell = "B"
    
        CheckAll
    
    Range("InitialsB2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsB2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

Sub Protect_InitialsB3()

Day_Date = "1"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsB3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsB3").Locked = False
    SelectedCell = "B"
    
    CheckAll
    
    Range("InitialsB3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsB3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsB3").Locked = False
    SelectedCell = "B"
    
    CheckAll
    
    Range("InitialsB3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsB3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


Sub Protect_InitialsC1()

Day_Date = "2"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsC1").Locked = False
    SelectedCell = "C"
    
    CheckAll
    
    Range("InitialsC1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsC1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsC1").Locked = False
    SelectedCell = "C"
    
    
    CheckAll
    
    
    Range("InitialsC1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsC1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub
    Sub Protect_InitialsC2()

Day_Date = "2"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsC2").Locked = False
    SelectedCell = "C"
    
    
    CheckAll
    
    
    Range("InitialsC2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsC2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsC2").Locked = False
    SelectedCell = "C"
    
    
    CheckAll
    
    
    Range("InitialsC2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsC2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If
    
End Sub

    Sub Protect_InitialsC3()

Day_Date = "2"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsC3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsC3").Locked = False
    SelectedCell = "C"
    
    
    CheckAll
    
    
    Range("InitialsC3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsC3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsC3").Locked = False
    SelectedCell = "C"
    
    
    CheckAll
    
    
    Range("InitialsC3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsC3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If
    
End Sub

    
        Sub Protect_InitialsD1()
        
        Day_Date = "3"
        
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsD1").Locked = False
    SelectedCell = "D"
    
    
    CheckAll
    
    
    Range("InitialsD1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsD1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsD1").Locked = False
    SelectedCell = "D"
    
    
    CheckAll
    
    
    Range("InitialsD1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsD1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


        Sub Protect_InitialsD2()

    Day_Date = "3"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsD2").Locked = False
    SelectedCell = "D"
    
    
    CheckAll
    
    
    Range("InitialsD2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsD2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsD2").Locked = False
    SelectedCell = "D"
    
    CheckAll
    
    Range("InitialsD2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsD2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsD3()

    Day_Date = "3"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsD3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsD3").Locked = False
    SelectedCell = "D"
    
    CheckAll
    
    Range("InitialsD3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsD3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsD3").Locked = False
    SelectedCell = "D"
    
    CheckAll
    
    Range("InitialsD3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsD3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


        Sub Protect_InitialsE1()

Day_Date = "4"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsE1").Locked = False
    SelectedCell = "E"
    
    CheckAll
    
    Range("InitialsE1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsE1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsE1").Locked = False
    SelectedCell = "E"
    
    CheckAll
    
    Range("InitialsE1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsE1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


        Sub Protect_InitialsE2()
        
        Day_Date = "4"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsE2").Locked = False
    SelectedCell = "E"
    
    CheckAll
    
    Range("InitialsE2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsE2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsE2").Locked = False
    SelectedCell = "E"
    
    CheckAll
    
    Range("InitialsE2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsE2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsE3()
        
        Day_Date = "4"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsE3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsE3").Locked = False
    SelectedCell = "E"
    
    CheckAll
    
    Range("InitialsE3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsE3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsE3").Locked = False
    SelectedCell = "E"
    
    CheckAll
    
    Range("InitialsE3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsE3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


        Sub Protect_InitialsF1()

Day_Date = "5"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsF1").Locked = False
    SelectedCell = "F"
    
    CheckAll
    
    Range("InitialsF1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsF1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsF1").Locked = False
    SelectedCell = "F"
    
    CheckAll
    
    Range("InitialsF1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsF1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


        Sub Protect_InitialsF2()

Day_Date = "5"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsF2").Locked = False
    SelectedCell = "F"
    
    CheckAll
    
    Range("InitialsF2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsF2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsF2").Locked = False
    SelectedCell = "F"
    
    CheckAll
    
    Range("InitialsF2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsF2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsF3()

Day_Date = "5"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsF3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsF3").Locked = False
    SelectedCell = "F"
    
    CheckAll
    
    Range("InitialsF3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsF3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsF3").Locked = False
    SelectedCell = "F"
    
    CheckAll
    
    Range("InitialsF3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsF3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


        Sub Protect_InitialsG1()
        
        Day_Date = "6"
        
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsG1").Locked = False
    SelectedCell = "G"
    
    CheckAll
    
    Range("InitialsG1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsG1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsG1").Locked = False
    SelectedCell = "G"
    
    CheckAll
    
    Range("InitialsG1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsG1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub



        Sub Protect_InitialsG2()

Day_Date = "6"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsG2").Locked = False
    SelectedCell = "G"
    
    CheckAll
    
    Range("InitialsG2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsG2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsG2").Locked = False
    SelectedCell = "G"
    
    CheckAll
    
    Range("InitialsG2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsG2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsG3()

Day_Date = "6"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsG3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsG3").Locked = False
    SelectedCell = "G"
    
    CheckAll
    
    Range("InitialsG3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsG3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsG3").Locked = False
    SelectedCell = "G"
    
    CheckAll
    
    Range("InitialsG3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsG3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


        Sub Protect_InitialsH1()

Day_Date = "7"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsH1").Locked = False
    SelectedCell = "H"
    
    CheckAll
    
    Range("InitialsH1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsH1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsH1").Locked = False
    SelectedCell = "H"
    
    CheckAll
    
    Range("InitialsH1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsH1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


        Sub Protect_InitialsH2()
        
        Day_Date = "7"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsH2").Locked = False
    SelectedCell = "H"
    
    CheckAll
    
    Range("InitialsH2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsH2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsH2").Locked = False
    SelectedCell = "H"
    
    CheckAll
    
    Range("InitialsH2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsH2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsH3()
        
        Day_Date = "7"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsH3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsH3").Locked = False
    SelectedCell = "H"
    
    CheckAll
    
    Range("InitialsH3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsH3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsH3").Locked = False
    SelectedCell = "H"
    
    CheckAll
    
    Range("InitialsH3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsH3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub




        Sub Protect_InitialsI1()

Day_Date = "8"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsI1").Locked = False
    SelectedCell = "I"
    
    CheckAll
    
    Range("InitialsI1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsI1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsI1").Locked = False
    SelectedCell = "I"
    
    CheckAll
    
    Range("InitialsI1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsI1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

        Sub Protect_InitialsI2()
        
        Day_Date = "8"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsI2").Locked = False
    SelectedCell = "I"
    
    CheckAll
    
    Range("InitialsI2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsI2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsI2").Locked = False
    SelectedCell = "I"
    
    CheckAll
    
    Range("InitialsI2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsI2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsI3()
        
        Day_Date = "8"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsI3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsI3").Locked = False
    SelectedCell = "I"
    
    CheckAll
    
    Range("InitialsI3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsI3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsI3").Locked = False
    SelectedCell = "I"
    
    CheckAll
    
    Range("InitialsI3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsI3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


        Sub Protect_InitialsJ1()

Day_Date = "9"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsJ1").Locked = False
    SelectedCell = "J"
    
    CheckAll
    
    Range("InitialsJ1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsJ1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsJ1").Locked = False
    SelectedCell = "J"
    
    CheckAll
    
    Range("InitialsJ1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsJ1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


        Sub Protect_InitialsJ2()
        
        Day_Date = "9"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsJ2").Locked = False
    SelectedCell = "J"
    
    CheckAll
    
    Range("InitialsJ2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsJ2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsJ2").Locked = False
    SelectedCell = "J"
    
    CheckAll
    
    Range("InitialsJ2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsJ2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsJ3()
        
        Day_Date = "9"
        
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsJ3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsJ3").Locked = False
    SelectedCell = "J"
    
    CheckAll
    
    Range("InitialsJ3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsJ3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsJ3").Locked = False
    SelectedCell = "J"
    
    CheckAll
    
    Range("InitialsJ3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsJ3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


        Sub Protect_InitialsK1()
        
        Day_Date = "10"
        
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsK1").Locked = False
    SelectedCell = "K"
    
    CheckAll
    
    Range("InitialsK1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsK1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsK1").Locked = False
    SelectedCell = "K"
    
    CheckAll
    
    Range("InitialsK1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsK1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


        Sub Protect_InitialsK2()

    Day_Date = "10"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsK2").Locked = False
    SelectedCell = "K"
    
    CheckAll
    
    Range("InitialsK2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsK2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsK2").Locked = False
    SelectedCell = "K"
    
    CheckAll
    
    Range("InitialsK2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsK2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsK3()

Day_Date = "10"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsK3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsK3").Locked = False
    SelectedCell = "K"
    
    CheckAll
    
    Range("InitialsK3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsK3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsK3").Locked = False
    SelectedCell = "K"
    
    CheckAll
    
    Range("InitialsK3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsK3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


        Sub Protect_InitialsL1()

Day_Date = "11"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsL1").Locked = False
    SelectedCell = "L"
    
    CheckAll
    
    Range("InitialsL1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsL1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsL1").Locked = False
    SelectedCell = "L"
    
    CheckAll
    
    Range("InitialsL1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsL1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

        Sub Protect_InitialsL2()

Day_Date = "11"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsL2").Locked = False
    SelectedCell = "L"
    
    CheckAll
    
    Range("InitialsL2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsL2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsL2").Locked = False
    SelectedCell = "L"
    
    CheckAll
    
    Range("InitialsL2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsL2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

        Sub Protect_InitialsL3()

Day_Date = "11"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsL3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsL3").Locked = False
    SelectedCell = "L"
    
    CheckAll
    
    Range("InitialsL3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsL3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsL3").Locked = False
    SelectedCell = "L"
    
    CheckAll
    
    Range("InitialsL3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsL3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


       Sub Protect_InitialsM1()

Day_Date = "12"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsM1").Locked = False
    SelectedCell = "M"
    
    CheckAll
    
    Range("InitialsM1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsM1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsM1").Locked = False
    SelectedCell = "M"
    
    CheckAll
    
    Range("InitialsM1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsM1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsM2()

Day_Date = "12"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsM2").Locked = False
    SelectedCell = "M"
    
    CheckAll
    
    Range("InitialsM2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsM2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsM2").Locked = False
    SelectedCell = "M"
    
    CheckAll
    
    Range("InitialsM2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsM2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsM3()

Day_Date = "12"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsM3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsM3").Locked = False
    SelectedCell = "M"
    
    CheckAll
    
    Range("InitialsM3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsM3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsM3").Locked = False
    SelectedCell = "M"
    
    CheckAll
    
    Range("InitialsM3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsM3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsN1()
       
       Day_Date = "13"
       
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsN1").Locked = False
    SelectedCell = "N"
    
    CheckAll
    
    Range("InitialsN1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsN1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsN1").Locked = False
    SelectedCell = "N"
    
    CheckAll
    
    Range("InitialsN1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsN1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsN2()

Day_Date = "13"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsN2").Locked = False
    SelectedCell = "N"
    
    CheckAll
    
    Range("InitialsN2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsN2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsN2").Locked = False
    SelectedCell = "N"
    
    CheckAll
    
    Range("InitialsN2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsN2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsN3()

Day_Date = "13"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsN3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsN3").Locked = False
    SelectedCell = "N"
    
    CheckAll
    
    Range("InitialsN3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsN3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsN3").Locked = False
    SelectedCell = "N"
    
    CheckAll
    
    Range("InitialsN3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsN3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsO1()
       
       Day_Date = "14"
       
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsO1").Locked = False
    SelectedCell = "O"
    
    CheckAll
    
    Range("InitialsO1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsO1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsO1").Locked = False
    SelectedCell = "O"
    
    CheckAll
    
    Range("InitialsO1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsO1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsO2()
    Day_Date = "14"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsO2").Locked = False
    SelectedCell = "O"
    
    CheckAll
    
    Range("InitialsO2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsO2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsO2").Locked = False
    SelectedCell = "O"
    
    CheckAll
    
    Range("InitialsO2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsO2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsO3()

Day_Date = "14"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsO3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsO3").Locked = False
    SelectedCell = "O"
    
    CheckAll
    
    Range("InitialsO3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsO3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsO3").Locked = False
    SelectedCell = "O"
    
    CheckAll
    
    Range("InitialsO3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsO3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsP1()
       
       Day_Date = "15"
       
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsP1").Locked = False
    SelectedCell = "P"
    
    CheckAll
    
    Range("InitialsP1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsP1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsP1").Locked = False
    SelectedCell = "P"
    
    CheckAll
    
    Range("InitialsP1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsP1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsP2()

    Day_Date = "15"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsP2").Locked = False
    SelectedCell = "P"
    
    CheckAll
    
    Range("InitialsP2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsP2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsP2").Locked = False
    SelectedCell = "P"
    
    CheckAll
    
    Range("InitialsP2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsP2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsP3()

Day_Date = "15"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsP3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsP3").Locked = False
    SelectedCell = "P"
    
    CheckAll
    
    Range("InitialsP3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsP3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsP3").Locked = False
    SelectedCell = "P"
    
    CheckAll
    
    Range("InitialsP3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsP3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsQ1()

Day_Date = "16"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsQ1").Locked = False
    SelectedCell = "Q"
    
    CheckAll
    
    Range("InitialsQ1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsQ1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsQ1").Locked = False
    SelectedCell = "Q"
    
    CheckAll
    
    Range("InitialsQ1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsQ1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsQ2()

Day_Date = "16"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsQ2").Locked = False
    SelectedCell = "Q"
    
    CheckAll
    
    Range("InitialsQ2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsQ2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsQ2").Locked = False
    SelectedCell = "Q"
    
    CheckAll
    
    Range("InitialsQ2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsQ2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsQ3()

Day_Date = "16"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsQ3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsQ3").Locked = False
    SelectedCell = "Q"
    
    CheckAll
    
    Range("InitialsQ3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsQ3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsQ3").Locked = False
    SelectedCell = "Q"
    
    CheckAll
    
    Range("InitialsQ3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsQ3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsR1()
       
       Day_Date = "17"
       
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsR1").Locked = False
    SelectedCell = "R"
    
    CheckAll
    
    Range("InitialsR1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsR1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsR1").Locked = False
    SelectedCell = "R"
    
    CheckAll
    
    Range("InitialsR1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsR1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsR2()

Day_Date = "17"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsR2").Locked = False
    SelectedCell = "R"
    
    CheckAll
    
    Range("InitialsR2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsR2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsR2").Locked = False
    SelectedCell = "R"
    
    CheckAll
    
    Range("InitialsR2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsR2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsR3()

Day_Date = "17"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsR3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsR3").Locked = False
    SelectedCell = "R"
    
    CheckAll
    
    Range("InitialsR3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsR3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsR3").Locked = False
    SelectedCell = "R"
    
    CheckAll
    
    Range("InitialsR3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsR3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsS1()
       
       Day_Date = "18"
       
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsS1").Locked = False
    SelectedCell = "S"
    
    CheckAll
    
    Range("InitialsS1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsS1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsS1").Locked = False
    SelectedCell = "S"
    
    CheckAll
    
    Range("InitialsS1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsS1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsS2()

Day_Date = "18"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsS2").Locked = False
    SelectedCell = "S"
    
    CheckAll
    
    Range("InitialsS2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsS2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsS2").Locked = False
    SelectedCell = "S"
    
    CheckAll
    
    Range("InitialsS2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsS2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If
    
End Sub

       Sub Protect_InitialsS3()

Day_Date = "18"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsS3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsS3").Locked = False
    SelectedCell = "S"
    
    CheckAll
    
    Range("InitialsS3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsS3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsS3").Locked = False
    SelectedCell = "S"
    
    CheckAll
    
    Range("InitialsS3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsS3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If
    
End Sub

       Sub Protect_InitialsT1()
       
       Day_Date = "19"
       
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsT1").Locked = False
    SelectedCell = "T"
    
    CheckAll
    
    Range("InitialsT1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsT1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsT1").Locked = False
    SelectedCell = "T"
    
    CheckAll
    
    Range("InitialsT1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsT1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


       Sub Protect_InitialsT2()
       
       Day_Date = "19"
       
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsT2").Locked = False
    SelectedCell = "T"
    
    CheckAll
    
    Range("InitialsT2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsT2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsT2").Locked = False
    SelectedCell = "T"
    
    CheckAll
    
    Range("InitialsT2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsT2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsT3()
       
       Day_Date = "19"
       
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsT3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsT3").Locked = False
    SelectedCell = "T"
    
    CheckAll
    
    Range("InitialsT3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsT3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsT3").Locked = False
    SelectedCell = "T"
    
    CheckAll
    
    Range("InitialsT3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsT3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


       Sub Protect_InitialsU1()
    
    Day_Date = "20"
        
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsU1").Locked = False
    SelectedCell = "U"
    
    CheckAll
    
    Range("InitialsU1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsU1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsU1").Locked = False
    SelectedCell = "U"
    
    CheckAll
    
    Range("InitialsU1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsU1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsU2()
       
       Day_Date = "20"
       
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsU2").Locked = False
    SelectedCell = "U"
    
    CheckAll
    
    Range("InitialsU2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsU2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsU2").Locked = False
    SelectedCell = "U"
    
    CheckAll
    
    Range("InitialsU2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsU2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsU3()
       
       Day_Date = "20"
       
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsU3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsU3").Locked = False
    SelectedCell = "U"
    
    CheckAll
    
    Range("InitialsU3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsU3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsU3").Locked = False
    SelectedCell = "U"
    
    CheckAll
    
    Range("InitialsU3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsU3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsV1()

       Day_Date = "21"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsV1").Locked = False
    SelectedCell = "V"
    
    CheckAll
    
    Range("InitialsV1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsV1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsV1").Locked = False
    SelectedCell = "V"
    
    CheckAll
    
    Range("InitialsV1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsV1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsV2()
       
       Day_Date = "21"
       
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsV2").Locked = False
    SelectedCell = "V"
    
    CheckAll
    
    Range("InitialsV2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsV2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsV2").Locked = False
    SelectedCell = "V"
    
    CheckAll
    
    Range("InitialsV2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsV2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsV3()
       
       Day_Date = "21"
       
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsV3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsV3").Locked = False
    SelectedCell = "V"
    
    CheckAll
    
    Range("InitialsV3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsV3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsV3").Locked = False
    SelectedCell = "V"
    
    CheckAll
    
    Range("InitialsV3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsV3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


       Sub Protect_InitialsW1()
       
       Day_Date = "22"
       
If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsW1").Locked = False
    SelectedCell = "W"
    
    CheckAll
    
    Range("InitialsW1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsW1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsW1").Locked = False
    SelectedCell = "W"
    
    CheckAll
    
    Range("InitialsW1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsW1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsW2()

    Day_Date = "22"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsW2").Locked = False
    SelectedCell = "W"
    
    CheckAll
    
    Range("InitialsW2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsW2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsW2").Locked = False
    SelectedCell = "W"
    
    CheckAll
    
    Range("InitialsW2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsW2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsW3()

    Day_Date = "22"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsW3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsW3").Locked = False
    SelectedCell = "W"
    
    CheckAll
    
    Range("InitialsW3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsW3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsW3").Locked = False
    SelectedCell = "W"
    
    CheckAll
    
    Range("InitialsW3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsW3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsX1()

Day_Date = "23"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsX1").Locked = False
    SelectedCell = "X"
    
    CheckAll
    
    Range("InitialsX1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsX1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsX1").Locked = False
    SelectedCell = "X"
    
    CheckAll
    
    Range("InitialsX1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsX1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsX2()

Day_Date = "23"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsX2").Locked = False
    SelectedCell = "X"
    
    CheckAll
    
    Range("InitialsX2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsX2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsX2").Locked = False
    SelectedCell = "X"
    
    CheckAll
    
    Range("InitialsX2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsX2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsX3()

Day_Date = "23"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsX3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsX3").Locked = False
    SelectedCell = "X"
    
    CheckAll
    
    Range("InitialsX3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsX3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsX3").Locked = False
    SelectedCell = "X"
    
    CheckAll
    
    Range("InitialsX3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsX3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsY1()

Day_Date = "24"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsY1").Locked = False
    SelectedCell = "Y"
    
    CheckAll
    
    Range("InitialsY1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsY1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsY1").Locked = False
    SelectedCell = "Y"
    
    CheckAll
    
    Range("InitialsY1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsY1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsY2()

Day_Date = "24"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsY2").Locked = False
    SelectedCell = "Y"
    
    CheckAll
    
    Range("InitialsY2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsY2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsY2").Locked = False
    SelectedCell = "Y"
    
    CheckAll
    
    Range("InitialsY2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsY2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsY3()

Day_Date = "24"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsY3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsY3").Locked = False
    SelectedCell = "Y"
    
    CheckAll
    
    Range("InitialsY3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsY3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsY3").Locked = False
    SelectedCell = "Y"
    
    CheckAll
    
    Range("InitialsY3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsY3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsZ1()

Day_Date = "25"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsZ1").Locked = False
    SelectedCell = "Z"
    
    CheckAll
    
    Range("InitialsZ1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsZ1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsZ1").Locked = False
    SelectedCell = "Z"
    
    CheckAll
    
    Range("InitialsZ1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsZ1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsZ2()

Day_Date = "25"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsZ2").Locked = False
    SelectedCell = "Z"
    
    CheckAll
    
    Range("InitialsZ2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsZ2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsZ2").Locked = False
    SelectedCell = "Z"
    
    CheckAll
    
    Range("InitialsZ2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsZ2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsZ3()

Day_Date = "25"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsZ3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsZ3").Locked = False
    SelectedCell = "Z"
    
    CheckAll
    
    Range("InitialsZ3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsZ3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsZ3").Locked = False
    SelectedCell = "Z"
    
    CheckAll
    
    Range("InitialsZ3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsZ3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


       Sub Protect_InitialsAA1()

Day_Date = "26"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAA1").Locked = False
    SelectedCell = "AA"
    
    CheckAll
    
    Range("InitialsAA1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAA1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAA1").Locked = False
    SelectedCell = "AA"
    
    CheckAll
    
    Range("InitialsAA1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAA1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsAA2()

Day_Date = "26"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAA2").Locked = False
    SelectedCell = "AA"
    
    CheckAll
    
    Range("InitialsAA2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAA2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAA2").Locked = False
    SelectedCell = "AA"
    
    CheckAll
    
    Range("InitialsAA2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAA2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAA3()

Day_Date = "26"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA3")) Then
MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAA3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAA3").Locked = False
    SelectedCell = "AA"
    
    CheckAll
    
    Range("InitialsAA3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAA3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAA3").Locked = False
    SelectedCell = "AA"
    
    CheckAll
    
    Range("InitialsAA3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAA3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAB1()

Day_Date = "27"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAB1").Locked = False
    SelectedCell = "AB"
    
    CheckAll
    
    Range("InitialsAB1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAB1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAB1").Locked = False
    SelectedCell = "AB"
    
    CheckAll
    
    Range("InitialsAB1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAB1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsAB2()

Day_Date = "27"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAB2").Locked = False
    SelectedCell = "AB"
    
    CheckAll
    
    Range("InitialsAB2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAB2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAB2").Locked = False
    SelectedCell = "AB"
    
    CheckAll
    
    Range("InitialsAB2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAB2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAB3()

Day_Date = "27"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAB3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAB3").Locked = False
    SelectedCell = "AB"
    
    CheckAll
    
    Range("InitialsAB3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAB3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAB3").Locked = False
    SelectedCell = "AB"
    
    CheckAll
    
    Range("InitialsAB3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAB3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAC1()

Day_Date = "28"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAC1").Locked = False
    SelectedCell = "AC"
    
    CheckAll
    
    Range("InitialsAC1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAC1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAC1").Locked = False
    SelectedCell = "AC"
    
    CheckAll
    
    Range("InitialsAC1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAC1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsAC2()

Day_Date = "28"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAC2").Locked = False
    SelectedCell = "AC"
    
    CheckAll
    
    Range("InitialsAC2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAC2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAC2").Locked = False
    SelectedCell = "AC"
    
    CheckAll
    
    Range("InitialsAC2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAC2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAC3()

Day_Date = "28"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAC3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAC3").Locked = False
    SelectedCell = "AC"
    
    CheckAll
    
    Range("InitialsAC3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAC3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAC3").Locked = False
    SelectedCell = "AC"
    
    CheckAll
    
    Range("InitialsAC3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAC3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAD1()

Day_Date = "29"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAD1").Locked = False
    SelectedCell = "AD"
    
    CheckAll
    
    Range("InitialsAD1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAD1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAD1").Locked = False
    SelectedCell = "AD"
    
    CheckAll
    
    Range("InitialsAD1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAD1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub


       Sub Protect_InitialsAD2()

Day_Date = "29"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAD2").Locked = False
    SelectedCell = "AD"
    
    CheckAll
    
    Range("InitialsAD2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAD2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAD2").Locked = False
    SelectedCell = "AD"
    
    CheckAll
    
    Range("InitialsAD2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAD2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAD3()

Day_Date = "29"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAD3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAD3").Locked = False
    SelectedCell = "AD"
    
    CheckAll
    
    Range("InitialsAD3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAD3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAD3").Locked = False
    SelectedCell = "AD"
    
    CheckAll
    
    Range("InitialsAD3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAD3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAE1()

Day_Date = "30"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAE1").Locked = False
    SelectedCell = "AE"
    
    CheckAll
    
    Range("InitialsAE1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAE1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAE1").Locked = False
    SelectedCell = "AE"
    
    CheckAll
    
    Range("InitialsAE1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAE1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsAE2()

Day_Date = "30"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAE2").Locked = False
    SelectedCell = "AE"
    
    CheckAll
    
    Range("InitialsAE2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAE2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAE2").Locked = False
    SelectedCell = "AE"
    
    CheckAll
    
    Range("InitialsAE2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAE2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAE3()

Day_Date = "30"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAE3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAE3").Locked = False
    SelectedCell = "AE"
    
    CheckAll
    
    Range("InitialsAE3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAE3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAE3").Locked = False
    SelectedCell = "AE"
    
    CheckAll
    
    Range("InitialsAE3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAE3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub


       Sub Protect_InitialsAF1()

Day_Date = "31"

If Application.WorksheetFunction.CountA(ThisWorkbook.Sheets(1).Range("TherapistName1:Title16")) > 0 Then
Else
MsgBox "Please enter your name, initials and title below before entering your initials in this cell."
Exit Sub
End If

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF1")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAF1").Locked = False
    SelectedCell = "AF"
    
    CheckAll
    
    Range("InitialsAF1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAF1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAF1").Locked = False
    SelectedCell = "AF"
    
    CheckAll
    
    Range("InitialsAF1").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAF1").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If

End Sub

       Sub Protect_InitialsAF2()

Day_Date = "31"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF2")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAF2").Locked = False
    SelectedCell = "AF"
    
    CheckAll
    
    Range("InitialsAF2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAF2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAF2").Locked = False
    SelectedCell = "AF"
    
    CheckAll
    
    Range("InitialsAF2").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAF2").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub

       Sub Protect_InitialsAF3()

Day_Date = "31"

If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF1")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF2")) And IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF3")) Then
    MsgBox ("Please use the above cell for initials first.")
    
    Exit Sub
    Else
If IsEmpty(ThisWorkbook.Sheets(1).Range("InitialsAF3")) Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAF3").Locked = False
    SelectedCell = "AF"
    
    CheckAll
    
    Range("InitialsAF3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAF3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
Dim Password As Variant
    Password = InputBox("Please enter password to re-enter initials", "Authenticate User", "")
    If Password = "healingarts" Then
    ActiveSheet.Unprotect "healingarts"
    ThisWorkbook.Sheets(1).Range("InitialsAF3").Locked = False
    SelectedCell = "AF"
    
    CheckAll
    
    Range("InitialsAF3").Value = Initials
    ThisWorkbook.Sheets(1).Range("InitialsAF3").Locked = True
    ActiveSheet.Protect "healingarts", AllowFormattingCells:=True
    
    Else
    MsgBox "Access denied - Please contact administrator", vbCritical, "Error"
    
    End If
    End If
    End If

End Sub





