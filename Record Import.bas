Attribute VB_Name = "Module11"
    Option Compare Text

Sub Import_Centre_Avenue_PT_Records()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim answer As Integer

    answer = MsgBox("Are you sure you want to import therapy records?", vbYesNo + vbQuestion, "Import Prompt")

    If answer = vbYes Then
    Else
    Exit Sub
    End If

   '''''Define Object for Target Workbook
    Dim Target_Workbook As Workbook
    Dim Source_Workbook As Workbook
    
    Set Source_Workbook = ThisWorkbook
    
    Source_Workbook.Sheets(2).Unprotect
    
    Dim lastrow As Long
    Dim lastrow_B As Long
    Dim lastrow_I As Long
    Dim lastrow_G As Long
    Dim lastrow_F As Long
    Dim lastrow_F2 As Long
    Dim path As String
    Dim Total_Files As Integer
    Dim minutes As Integer
    Dim discipline As String
    
    Dim therapy_begin_date As Date
    Dim therapy_end_date As Date
    
    Dim Target_Folder As String
    
    Dim StrFile As String
    
    Target_Folder = Source_Workbook.Sheets(1).Range("MonthYear").Value
    
    StrFile = Dir("G:\Therapy Charting Grids\Centre Avenue\Therapy Charts\" & Target_Folder & "\PT\*")
    
    Do While Len(StrFile) > 0
    
    Application.EnableEvents = False
    Set Target_Workbook = Workbooks.Open("G:\Therapy Charting Grids\Centre Avenue\Therapy Charts\" & Target_Folder & "\PT\" & StrFile)
    Application.EnableEvents = True
    
    path = "G:\Therapy Charting Grids\Centre Avenue\Therapy Charts\" & Target_Folder & "\PT\" & StrFile
    
    Dim Month As String
    Dim Day As String
    Dim Year As String
    Dim Quantity As Integer
    Dim Lastname1 As String
    Dim Lastname2 As String
    Dim FirstName As String
    Dim RowNumber As String
    Dim i As Range
    Dim TestArray() As String
    Dim MRN As String
    Dim Range_Test As Range
    Dim Admission As Variant
    Dim Discharge As Variant
    
    If InStr(1, (Target_Workbook.Name), "Jan ") > 0 Then
    Month = "01"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Feb ") > 0 Then
    Month = "02"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Mar ") > 0 Then
    Month = "03"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Apr ") > 0 Then
    Month = "04"
    End If
    
    If InStr(1, (Target_Workbook.Name), "May ") > 0 Then
    Month = "05"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Jun ") > 0 Then
    Month = "06"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Jul ") > 0 Then
    Month = "07"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Aug ") > 0 Then
    Month = "08"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Sep ") > 0 Then
    Month = "09"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Oct ") > 0 Then
    Month = "10"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Nov ") > 0 Then
    Month = "11"
    End If
    
    If InStr(1, (Target_Workbook.Name), "Dec ") > 0 Then
    Month = "12"
    End If
    
    If InStr(1, (Target_Workbook.Name), "16") > 0 Then
    Year = "2016"
    End If
    
    If InStr(1, (Target_Workbook.Name), "17") > 0 Then
    Year = "2017"
    End If
    
    If InStr(1, (Target_Workbook.Name), "18") > 0 Then
    Year = "2018"
    End If
    
    Quantity = Target_Workbook.Sheets(1).Range("CurrentMonthVisits").Value
    
        For Each i In Target_Workbook.ActiveSheet.Range("AG1:AG100")
        If InStr(1, Target_Workbook.ActiveSheet.Range("AG" & i.Row).Value, "Total Mins ") Then
            Target_Workbook.Sheets(1).Range("AG" & i.Row).Offset(rowOffset:=1).Activate
            minutes = ActiveCell.Value
        End If
        Next
    
    '''''With Target_Workbook object now, it is possible to pull any data from it
    '''''Read Data from Target File
        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsB1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsB2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsB3")) Then
        Else
        Day = "01"
        GoTo Step2
        End If
            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsC1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsC2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsB3")) Then
            Else
            Day = "02"
            GoTo Step2
            End If
                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsD1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsD2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsD3")) Then
                Else
                Day = "03"
                GoTo Step2
                End If
                    If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsE1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsE2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsE3")) Then
                    Else
                    Day = "04"
                    GoTo Step2
                    End If
                        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsF1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsF2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsF3")) Then
                        Else
                        Day = "05"
                        GoTo Step2
                        End If
                            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsG1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsG2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsG3")) Then
                            Else
                            Day = "06"
                            GoTo Step2
                            End If
                                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsH1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsH2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsH3")) Then
                                Else
                                Day = "07"
                                GoTo Step2
                                End If
                                    If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsI1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsI2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsI3")) Then
                                    Else
                                    Day = "08"
                                    GoTo Step2
                                    End If
                                        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsJ1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsJ2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsJ3")) Then
                                        Else
                                        Day = "09"
                                        GoTo Step2
                                        End If
                                            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsK1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsK2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsK3")) Then
                                            Else
                                            Day = "10"
                                            GoTo Step2
                                            End If
                                                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsL1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsL2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsL3")) Then
                                                Else
                                                Day = "11"
                                                GoTo Step2
                                                End If
                                                    If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsM1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsM2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsM3")) Then
                                                    Else
                                                    Day = "12"
                                                    GoTo Step2
                                                    End If
                                                        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsN1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsN2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsN3")) Then
                                                        Else
                                                        Day = "13"
                                                        GoTo Step2
                                                        End If
                                                            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsO1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsO2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsO3")) Then
                                                            Else
                                                            Day = "14"
                                                            GoTo Step2
                                                            End If
                                                                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsP1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsP2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsP3")) Then
                                                                Else
                                                                Day = "15"
                                                                GoTo Step2
                                                                End If
                                                                    If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsQ1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsQ2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsQ3")) Then
                                                                    Else
                                                                    Day = "16"
                                                                    GoTo Step2
                                                                    End If
                                                                        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsR1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsR2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsR3")) Then
                                                                        Else
                                                                        Day = "17"
                                                                        GoTo Step2
                                                                        End If
                                                                            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsS1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsS2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsS3")) Then
                                                                            Else
                                                                            Day = "18"
                                                                            GoTo Step2
                                                                            End If
                                                                                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsT1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsT2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsT3")) Then
                                                                                Else
                                                                                Day = "19"
                                                                                GoTo Step2
                                                                                End If
                                                                                    If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsU1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsU2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsU3")) Then
                                                                                    Else
                                                                                    Day = "20"
                                                                                    GoTo Step2
                                                                                    End If
                                                                                        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsV1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsV2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsV3")) Then
                                                                                        Else
                                                                                        Day = "21"
                                                                                        GoTo Step2
                                                                                        End If
                                                                                            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsW1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsW2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsW3")) Then
                                                                                            Else
                                                                                            Day = "22"
                                                                                            GoTo Step2
                                                                                            End If
                                                                                                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsX1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsX2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsX3")) Then
                                                                                                Else
                                                                                                Day = "23"
                                                                                                GoTo Step2
                                                                                                End If
                                                                                                    If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsY1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsY2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsY3")) Then
                                                                                                    Else
                                                                                                    Day = "24"
                                                                                                    GoTo Step2
                                                                                                    End If
                                                                                                        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsZ1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsZ2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsZ3")) Then
                                                                                                        Else
                                                                                                        Day = "25"
                                                                                                        GoTo Step2
                                                                                                        End If
                                                                                                            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAA1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAA2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAA3")) Then
                                                                                                            Else
                                                                                                            Day = "26"
                                                                                                            GoTo Step2
                                                                                                            End If
                                                                                                                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAB1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAB2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAB3")) Then
                                                                                                                Else
                                                                                                                Day = "27"
                                                                                                                GoTo Step2
                                                                                                                End If
                                                                                                                    If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAC1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAC2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAC3")) Then
                                                                                                                    Else
                                                                                                                    Day = "28"
                                                                                                                    GoTo Step2
                                                                                                                    End If
                                                                                                                        If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAD1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAD2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAD3")) Then
                                                                                                                        Else
                                                                                                                        Day = "29"
                                                                                                                        GoTo Step2
                                                                                                                        End If
                                                                                                                            If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAE1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAE2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAE3")) Then
                                                                                                                            Else
                                                                                                                            Day = "30"
                                                                                                                            GoTo Step2
                                                                                                                            End If
                                                                                                                                If IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAF1")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAF2")) And IsEmpty(Target_Workbook.Sheets(1).Range("InitialsAF3")) Then
                                                                                                                                Else
                                                                                                                                Day = "31"
                                                                                                                                GoTo Step2
                                                                                                                                End If

Step2:

            If InStr(1, (path), "\Centre Avenue\") > 0 Then
                lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "J").End(xlUp).Row + 1
                Source_Workbook.Sheets(2).Cells(lastrow, "J").Value = "8"
                
                lastrow_B = Source_Workbook.Sheets(2).Cells(Rows.count, "K").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_B, "K").Value = "CENTRE AVENUE HEALTH & REHAB - Ft. Collins, CO"
            End If
            
            If InStr(1, (path), "\Columbine Commons\") > 0 Then
                lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "J").End(xlUp).Row + 1
                Source_Workbook.Sheets(2).Cells(lastrow, "J").Value = "27"
                
                lastrow_B = Source_Workbook.Sheets(2).Cells(Rows.count, "K").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_B, "K").Value = "COLUMBINE COMMONS HEALTH & REHAB - CO"
            End If
            
            If InStr(1, (path), "\Lemay Avenue\") > 0 Then
                lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "J").End(xlUp).Row + 1
                Source_Workbook.Sheets(2).Cells(lastrow, "J").Value = "13"
                
                lastrow_B = Source_Workbook.Sheets(2).Cells(Rows.count, "K").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_B, "K").Value = "LEMAY AVENUE HEALTH & REHAB - Fort Collins, CO"
            End If
            
            If InStr(1, (path), "\Columbine West\") > 0 Then
                lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "J").End(xlUp).Row + 1
                Source_Workbook.Sheets(2).Cells(lastrow, "J").Value = "14"
                
                lastrow_B = Source_Workbook.Sheets(2).Cells(Rows.count, "K").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_B, "K").Value = "COLUMBINE WEST HEALTH & REHAB - Fort Collins, CO"
            End If
            
            If InStr(1, (path), "\North Shore\") > 0 Then
                lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "J").End(xlUp).Row + 1
                Source_Workbook.Sheets(2).Cells(lastrow, "J").Value = "15"
                
                lastrow_B = Source_Workbook.Sheets(2).Cells(Rows.count, "K").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_B, "K").Value = "NORTH SHORE HEALTH & REHAB - Loveland, CO"
            End If
            
            lastrow_I = Source_Workbook.Sheets(2).Cells(Rows.count, "R").End(xlUp).Row
            Source_Workbook.Sheets(2).Cells(lastrow_I, "R").Value = Month & "-" & Day & "-" & Year
            
            If InStr(1, (path), "\OT\") > 0 Then
                lastrow_G = Source_Workbook.Sheets(2).Cells(Rows.count, "P").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_G, "P").Value = "OT"
                discipline = "OT"
            End If
            
            If InStr(1, (path), "\PT\") > 0 Then
                lastrow_G = Source_Workbook.Sheets(2).Cells(Rows.count, "P").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_G, "P").Value = "PT"
                discipline = "PT"
            End If
            
            If InStr(1, (path), "\ST\") > 0 Then
                lastrow_G = Source_Workbook.Sheets(2).Cells(Rows.count, "P").End(xlUp).Row
                Source_Workbook.Sheets(2).Cells(lastrow_G, "P").Value = "ST"
                discipline = "ST"
            End If
            
            lastrow_K = Source_Workbook.Sheets(2).Cells(Rows.count, "T").End(xlUp).Row
            Source_Workbook.Sheets(2).Cells(lastrow_K, "T").Value = Quantity
            
            lastrow_K = Source_Workbook.Sheets(2).Cells(Rows.count, "H").End(xlUp).Row + 1
            Source_Workbook.Sheets(2).Cells(lastrow_K, "H").Value = Quantity
            


            lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "I").End(xlUp).Row + 1
            Source_Workbook.Sheets(2).Cells(lastrow, "I").Value = minutes
            
            lastrow_I = Source_Workbook.Sheets(2).Cells(Rows.count, "D").End(xlUp).Row + 1
            therapy_begin_date = Month & "-" & Day & "-" & Year
            Source_Workbook.Sheets(2).Cells(lastrow_I, "D").Value = therapy_begin_date
   
        If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesB").Value > 0 Then
        
        Day = "01"
        End If
            If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesC").Value > 0 Then
            
            Day = "02"
            End If
                If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesD").Value > 0 Then
                
                Day = "03"
                End If
                    If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesE").Value > 0 Then
                    
                    Day = "04"
                    End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesF").Value > 0 Then

                        Day = "05"
                        End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesG").Value > 0 Then

                            Day = "06"
                            End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesH").Value > 0 Then

                                Day = "07"
                                End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesI").Value > 0 Then

                                    Day = "08"
                                    End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesJ").Value > 0 Then

                                        Day = "09"
                                        End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesK").Value > 0 Then

                                            Day = "10"
                                            End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesL").Value > 0 Then

                                                Day = "11"
                                                End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesM").Value > 0 Then

                                                    Day = "12"
                                                    End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesN").Value > 0 Then

                                                        Day = "13"
                                                        End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesO").Value > 0 Then

                                                            Day = "14"
                                                            End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesP").Value > 0 Then

                                                                Day = "15"
                                                                End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesQ").Value > 0 Then

                                                                    Day = "16"
                                                                    End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesR").Value > 0 Then

                                                                        Day = "17"
                                                                        End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesS").Value > 0 Then

                                                                            Day = "18"
                                                                            End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesT").Value > 0 Then

                                                                                Day = "19"
                                                                                End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesU").Value > 0 Then

                                                                                    Day = "20"
                                                                                    End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesV").Value > 0 Then

                                                                                        Day = "21"
                                                                                        End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesW").Value > 0 Then

                                                                                            Day = "22"
                                                                                            End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesX").Value > 0 Then

                                                                                                Day = "23"
                                                                                                End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesY").Value > 0 Then

                                                                                                    Day = "24"
                                                                                                    End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesZ").Value > 0 Then

                                                                                                        Day = "25"
                                                                                                        End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesAA").Value > 0 Then

                                                                                                            Day = "26"
                                                                                                            End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesAB").Value > 0 Then

                                                                                                                Day = "27"
                                                                                                                End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesAC").Value > 0 Then

                                                                                                                    Day = "28"
                                                                                                                    End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesAD").Value > 0 Then

                                                                                                                        Day = "29"
                                                                                                                        End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesAE").Value > 0 Then

                                                                                                                            Day = "30"
                                                                                                                            End If
If Target_Workbook.Sheets(1).Range("TotalTreatmentMinutesAF").Value > 0 Then

                                                                                                                                Day = "31"
                                                                                                                                End If
            
            lastrow_I = Source_Workbook.Sheets(2).Cells(Rows.count, "E").End(xlUp).Row + 1
            therapy_end_date = Month & "/" & Day & "/" & Year
            Source_Workbook.Sheets(2).Cells(lastrow_I, "E").Value = therapy_end_date
            
    Lastname1 = Target_Workbook.Sheets(1).Range("H3").Value
    FirstName = Target_Workbook.Sheets(1).Range("O3").Value
    
    Dim FirstName1 As String
    Dim FirstName2 As String
    Dim admission_switch As Boolean
    Dim discharge_switch As Boolean
    Dim return_switch As Boolean
    Dim new_visit_switch As Boolean
    Dim target_name As String
    Dim mrn_array() As String
    Dim admission_2 As Variant
    Dim discharge_2 As Variant
    
    admission_switch = False
    discharge_switch = False
    return_switch = False
    new_visit_switch = True
    
    target_name = Target_Workbook.Name
            
    If InStr(1, target_name, " V 1") Or InStr(1, target_name, " V 2") Or InStr(1, target_name, " V 3") Or InStr(1, target_name, " V 4") Then
        new_visit_switch = False
    End If
    
'    On Error GoTo MissingMRNSheet
    
    With Source_Workbook.Sheets(3)
    
    For Each i In .Range("B1:B1000")
    
    
    If .Range("A" & i.Row).Value = "Admission" Then
        admission_switch = True
        discharge_switch = False
        return_switch = False
    End If
    
    If .Range("A" & i.Row).Value = "Discharged Return Not Anticipated" Then
        admission_switch = False
        discharge_switch = True
        return_switch = False
    End If
        
    If .Range("A" & i.Row).Value = "Discharged Return Expected" Then
        admission_switch = False
        discharge_switch = True
        return_switch = False
    End If
    
    If .Range("A" & i.Row).Value = "Return" Then
        admission_switch = False
        discharge_switch = False
        return_switch = True
    End If
    
        
    If .Range("A" & i.Row).Value = "Expired" Then
        admission_switch = False
        discharge_switch = True
        return_switch = False
    End If
    
    If .Range("A" & i.Row).Value = "Hospital Leave" Then
        admission_switch = False
        discharge_switch = False
        return_switch = False
    End If
    
    If .Range("A" & i.Row).Value = "Outpatient" Then
        admission_switch = False
        discharge_switch = False
        return_switch = False
    End If
    
    If .Range("A" & i.Row).Value = "Therapeutic Leave" Then
        admission_switch = False
        discharge_switch = False
        return_switch = False
    End If
    
    If InStr(1, .Range("B" & i.Row).Value, ",") Then
    TestArray() = Split(.Range("B" & i.Row).Value, ", ")
    FirstName2 = Left(TestArray(1), 2)
    FirstName1 = Left(FirstName, 2)
    FirstName2 = Trim(FirstName2)
    FirstName1 = Trim(FirstName1)
    Lastname2 = TestArray(0)
    If Lastname1 = Lastname2 And FirstName1 = FirstName2 Then

    mrn_array() = Split(.Range("E" & i.Row).Value, "-")
    MRN = mrn_array(0)

    If admission_switch = True And new_visit_switch = True Then
        Admission = .Range("A" & i.Row).Value
        Admission = CDate(CLng(Admission)) - 1
    End If
    
    If return_switch = True Then
        admission_2 = .Range("A" & i.Row).Value
        admission_2 = CDate(CLng(admission_2)) - 1
        If admission_2 <= therapy_begin_date Then
            Admission = admission_2
        End If
    End If
    
    If discharge_switch = True Then
        
        discharge_2 = .Range("A" & i.Row).Value

        discharge_2 = CDate(CLng(discharge_2))
        If discharge_2 >= therapy_end_date Then
            Discharge = discharge_2
        End If
        
    End If
    
    End If
    End If
    Next i
    
    On Error GoTo 0
    
    End With
    
            
            lastrow_F = Source_Workbook.Sheets(2).Cells(Rows.count, "O").End(xlUp).Row
            Application.EnableEvents = False
            Source_Workbook.Sheets(2).Cells(lastrow_F, "O").Value = MRN
            Application.EnableEvents = True
            
            If Admission = "" Then
            Admission = "Not Found"
            End If
            
            If Discharge = "" Then
            Discharge = "Not Found"
            End If
            
            'If Not Admission = therapy_begin_date Then
            
            Dim visit_number As String
            visit_number = ""
            
            If InStr(1, target_name, " V 1") Or InStr(1, target_name, " V 2") Or InStr(1, target_name, " V 3") Or InStr(1, target_name, " V 4") Then
                visit_number = "(Return Visit)"
            End If
            
            lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "C").End(xlUp).Row + 1
            Source_Workbook.Sheets(2).Cells(lastrow, "C").Value = Lastname1 & ", " & FirstName & " " & "(" & discipline & Total_Files & ")" & " " & visit_number
            
            lastrow = Source_Workbook.Sheets(2).Cells(Rows.count, "N").End(xlUp).Row

            Source_Workbook.Sheets(2).Cells(lastrow, "N").Value = Lastname1 & ", " & FirstName & " " & "(" & discipline & Total_Files & ")" & " " & visit_number
            
            lastrow_F = Source_Workbook.Sheets(2).Cells(Rows.count, "F").End(xlUp).Row + 1
            Source_Workbook.Sheets(2).Cells(lastrow_F, "F").Value = Admission
            
            lastrow_I = Source_Workbook.Sheets(2).Cells(Rows.count, "G").End(xlUp).Row + 1
            Source_Workbook.Sheets(2).Cells(lastrow_I, "G").Value = Discharge
            
            If MRN = "" Then
            'Set Range_Test = Source_Workbook.Sheets(2).Cells(Rows.count, "I").End(xlUp)
            
            'lastrow = Source_Workbook.Sheets(1).Cells(Rows.count, "M").End(xlUp).Row + 1
            'Source_Workbook.Sheets(1).Cells(lastrow, "M") = Range_Test.Row
            
            'lastrow = Source_Workbook.Sheets(1).Cells(Rows.count, "N").End(xlUp).Row + 1
            'Source_Workbook.Sheets(1).Cells(lastrow, "N").Value = Lastname1 & ", " & FirstName
            
            End If
            
            Source_Workbook.Sheets(2).Unprotect
            
    '''''Close Target Workbook
    Application.EnableEvents = False
    Target_Workbook.Close
    Application.EnableEvents = True
    
    MRN = ""
    Admission = ""
    Discharge = ""
    discharge_2 = ""
    admission_2 = ""
    
    Total_Files = Total_Files + 1
    
    StrFile = Dir

    Loop
       
    Source_Workbook.Sheets(2).Range("B3").Value = Total_Files
    
    Dim FolderPath As String, path2 As String, count As Integer
    FolderPath = "G:\Therapy Charting Grids\Centre Avenue\Therapy Charts\" & Target_Folder & "\PT\"

    path2 = FolderPath & "\*.xlsm"

    Filename = Dir(path2)

    Do While Filename <> ""
        count = count + 1
        Filename = Dir()
    Loop

    Source_Workbook.Sheets(2).Range("B6") = count
    'MsgBox count & " : files found in folder"

    Source_Workbook.Sheets(2).Protect
    
    Source_Workbook.Save

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub



