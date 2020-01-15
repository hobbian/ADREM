Attribute VB_Name = "Sample_Totals_Calculations"
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine Written by Ian M. Hobbs 26 March 2018
'Subroutine Transfers the data from Raw Data Tab to the Sample Totals tab
    'For further refinement by dilution factors
    
Sub Transfer_Data()
Attribute Transfer_Data.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Integer
Dim RNG As Range
Dim ViewMode As Long
Dim CalcMode As Long

With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
End With

    'Delete Non Empricial Data From Row E
        'Looks for data in worksheet labled Raw Data
        With Sheets("Sample Totals")
        'Selects the sheet
            .Select
        'Changes the view to normal view
            ViewMode = ActiveWindow.View
            ActiveWindow.View = xlNormalView
        'Turn off Page Braks for Speed
            .DisplayPageBreaks = False
            
On Error GoTo Error_Handler_TD:

i = 2

    Sheets("Raw Data").Select
        Columns("E:E").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("E1").Select
                ActiveCell.FormulaR1C1 = "HelperTab"
                    
        Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))

            Cells(i, "E").FormulaR1C1 = "=RC[-4]&RC[-3]&RC[-1]"
                
            i = i + 1
        
        Loop
        
i = 27

    Sheets("Sample Totals").Select
    
        'Stops the function once the data set is complete
        Do While Not IsEmpty(Sheets("Sample Totals").Cells(i, "A"))
                            
            'Inputs data corresponding to the AL# the Sample ID and the Element/Isotope
            'and paste it on Sample Totals Tab
            Cells(i, "H") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "I") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "J") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "K") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "L") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "M") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "N") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "O") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "P") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "Q") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "R") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "S") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "T") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "U") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "V") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "W") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "X") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "Y") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "Z") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AA") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AB") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AC") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AD") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AE") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AF") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AG") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AH") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AI") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AJ") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AK") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AL") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AM") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AN") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AO") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AP") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AQ") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AR") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AS") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AT") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AU") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AV") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AW") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AX") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AY") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "AZ") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "BA") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
            Cells(i, "BB") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R24C,'Raw Data'!C5:C6,2,FALSE)"
                              
                
        i = i + 1
        
    Loop
    
i = 27

    Sheets("Sample Totals").Select
    
        'Determines the range in which the data lies
        Lrow = Range("H" & Rows.Count).End(xlUp).Row
        Lcol = Cells(27, Columns.Count).End(xlToLeft).Column
    
            Set RNG = Range(Cells(27, "H"), Cells(Lrow, Lcol))
              
                'Sets the output to scientific notation and changes font size and style
                With RNG
                    .Copy
                End With
                    Cells(27, "H").PasteSpecial xlPasteValues
                    
        Sheets("Raw Data").Select
        Columns("E:E").Delete
        Sheets("Sample Totals").Select
        RNG.ClearFormats
    
    End With
    
Exit Sub

Error_Handler_TD:
    MsgBox " Incomplete Data set insure all data has been entered and processed. :1"
    Sheets("Raw Data").Select
    Columns("E:E").Delete
    Sheets("Sample Totals").Select
    RNG.ClearFormats
    
End Sub
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine Writtern by Ian M. Hobbs 29 March 2018
'Subroutine calculates the total grams of analyte disposed
    ' in the consolidation
    
Sub Sample_Totals()
Dim i As Integer
Dim RNG As Range
Dim Lrow As Long
Dim Lcol As Long

i = 27
    
    'Selects the appropriate sheet to perform calculation
    Sheets("Sample Totals").Select
    
            'Performs calulation as long as there are AL number in column A
            Do While Not IsEmpty(Sheets("Sample Totals").Cells(i, "A"))
                       
              
                               
                    'Determines if cell containes a numerical value
                    'Performs diultion calucation based on input values
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "H").Interior.Color = RGB(160, 160, 530) Or Cells(i, "H").Interior.Color = RGB(0, 150, 520) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        
                        On Error GoTo Error_Handler:
                        
                        If IsNumeric(Cells(i, "H")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "H") = Cells(i, "H") * Cells(i, "F")
                                Cells(i, "H").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "H") = (Cells(i, "H") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "H").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                        
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "I").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        
                        On Error GoTo Error_Handler:
                        
                        If IsNumeric(Cells(i, "I")) Then
                           If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "I") = Cells(i, "I") * Cells(i, "F")
                                Cells(i, "I").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "I") = (Cells(i, "I") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "I").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                        
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "J").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "J")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "J") = Cells(i, "J") * Cells(i, "F")
                                Cells(i, "J").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "J") = (Cells(i, "J") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "J").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "K").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "K")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "K") = Cells(i, "K") * Cells(i, "F")
                                Cells(i, "K").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "K") = (Cells(i, "K") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "K").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                                             
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "L").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "L")) Then
                              If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "L") = Cells(i, "L") * Cells(i, "F")
                                Cells(i, "L").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "L") = (Cells(i, "L") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "L").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "M").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "M")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "M") = Cells(i, "M") * Cells(i, "F")
                                Cells(i, "M").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "M") = (Cells(i, "M") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "M").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "N").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "N")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "N") = Cells(i, "N") * Cells(i, "F")
                                Cells(i, "N").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "N") = (Cells(i, "N") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "N").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "O").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "O")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "O") = Cells(i, "O") * Cells(i, "F")
                                Cells(i, "O").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "O") = (Cells(i, "O") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "O").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "P").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "P")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "P") = Cells(i, "P") * Cells(i, "F")
                                Cells(i, "P").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "P") = (Cells(i, "P") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "P").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "Q").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "Q")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "Q") = Cells(i, "Q") * Cells(i, "F")
                                Cells(i, "Q").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "Q") = (Cells(i, "Q") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "Q").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "R").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "R")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "R") = Cells(i, "R") * Cells(i, "F")
                                Cells(i, "R").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "R") = (Cells(i, "R") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "R").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "S").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "S")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "S") = Cells(i, "S") * Cells(i, "F")
                                Cells(i, "S").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "S") = (Cells(i, "S") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "S").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "T").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "T")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "T") = Cells(i, "T") * Cells(i, "F")
                                Cells(i, "T").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "T") = (Cells(i, "T") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "T").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "U").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "U")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "U") = Cells(i, "U") * Cells(i, "F")
                                Cells(i, "U").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "U") = (Cells(i, "U") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "U").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "V").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "V")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "V") = Cells(i, "V") * Cells(i, "F")
                                Cells(i, "V").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "V") = (Cells(i, "V") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "V").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "W").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "W")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "W") = Cells(i, "W") * Cells(i, "F")
                                Cells(i, "W").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "W") = (Cells(i, "W") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "W").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "X").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "X")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "X") = Cells(i, "X") * Cells(i, "F")
                                Cells(i, "X").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "X") = (Cells(i, "X") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "X").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "Y").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "Y")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "Y") = Cells(i, "Y") * Cells(i, "F")
                                Cells(i, "Y").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "Y") = (Cells(i, "Y") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "Y").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "Z").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "Z")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "Z") = Cells(i, "Z") * Cells(i, "F")
                                Cells(i, "Z").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "Z") = (Cells(i, "Z") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "Z").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AA").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AA")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AA") = Cells(i, "AA") * Cells(i, "F")
                                Cells(i, "AA").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AA") = (Cells(i, "AA") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AA").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AB").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AB")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AB") = Cells(i, "AB") * Cells(i, "F")
                                Cells(i, "AB").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AB") = (Cells(i, "AB") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AB").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AC").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AC")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AC") = Cells(i, "AC") * Cells(i, "F")
                                Cells(i, "AC").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AC") = (Cells(i, "AC") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AC").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AD").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AD")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AD") = Cells(i, "AD") * Cells(i, "F")
                                Cells(i, "AD").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AD") = (Cells(i, "AD") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AD").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AE").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AE")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AE") = Cells(i, "AE") * Cells(i, "F")
                                Cells(i, "AE").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AE") = (Cells(i, "AE") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AE").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AF").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AF")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AF") = Cells(i, "AF") * Cells(i, "F")
                                Cells(i, "AF").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AF") = (Cells(i, "AF") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AF").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AG").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AG")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AG") = Cells(i, "AG") * Cells(i, "F")
                                Cells(i, "AG").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AG") = (Cells(i, "AG") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AG").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AH").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AH")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AH") = Cells(i, "AH") * Cells(i, "F")
                                Cells(i, "AH").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AH") = (Cells(i, "AH") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AH").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AI").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AI")) Then
                           If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AI") = Cells(i, "AI") * Cells(i, "F")
                                Cells(i, "AI").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AI") = (Cells(i, "AI") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AI").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AJ").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AJ")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AJ") = Cells(i, "AJ") * Cells(i, "F")
                                Cells(i, "AJ").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AJ") = (Cells(i, "AJ") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AJ").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AK").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AK")) Then
                           If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AK") = Cells(i, "AK") * Cells(i, "F")
                                Cells(i, "AK").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AK") = (Cells(i, "AK") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AK").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                
                    End If
                
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AL").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AL")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AL") = Cells(i, "AL") * Cells(i, "F")
                                Cells(i, "AL").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AL") = (Cells(i, "AL") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AL").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                                  
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AM").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AM")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AM") = Cells(i, "AM") * Cells(i, "F")
                                Cells(i, "AM").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AM") = (Cells(i, "AM") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AM").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AN").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AN")) Then
                           If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AN") = Cells(i, "AN") * Cells(i, "F")
                                Cells(i, "AN").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AN") = (Cells(i, "AN") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AN").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AO").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AO")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AO") = Cells(i, "AO") * Cells(i, "F")
                                Cells(i, "AO").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AO") = (Cells(i, "AO") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AO").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AP").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AP")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AP") = Cells(i, "AP") * Cells(i, "F")
                                Cells(i, "AP").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AP") = (Cells(i, "AP") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AP").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AQ").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AQ")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AQ") = Cells(i, "AQ") * Cells(i, "F")
                                Cells(i, "AQ").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AQ") = (Cells(i, "AQ") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AQ").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AR").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AR")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AR") = Cells(i, "AR") * Cells(i, "F")
                                Cells(i, "AR").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AR") = (Cells(i, "AR") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AR").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AS").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AS")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AS") = Cells(i, "AS") * Cells(i, "F")
                                Cells(i, "AS").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AS") = (Cells(i, "AS") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AS").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AT").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AT")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AT") = Cells(i, "AT") * Cells(i, "F")
                                Cells(i, "AT").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AT") = (Cells(i, "AT") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AT").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AU").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AU")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AU") = Cells(i, "AU") * Cells(i, "F")
                                Cells(i, "AU").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AU") = (Cells(i, "AU") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AU").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AV").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AV")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AV") = Cells(i, "AV") * Cells(i, "F")
                                Cells(i, "AV").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AV") = (Cells(i, "AV") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AV").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AW").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AW")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AW") = Cells(i, "AW") * Cells(i, "F")
                                Cells(i, "AW").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AW") = (Cells(i, "AW") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AW").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AX").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AX")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AX") = Cells(i, "AX") * Cells(i, "F")
                                Cells(i, "AX").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AX") = (Cells(i, "AX") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AX").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AY").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AY")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AY") = Cells(i, "AY") * Cells(i, "F")
                                Cells(i, "AY").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AY") = (Cells(i, "AY") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AY").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "AZ").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "AZ")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "AZ") = Cells(i, "AZ") * Cells(i, "F")
                                Cells(i, "AZ").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "AZ") = (Cells(i, "AZ") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "AZ").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "BA").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "BA")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "BA") = Cells(i, "BA") * Cells(i, "F")
                                Cells(i, "BA").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "BA") = (Cells(i, "BA") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "BA").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
                    If Not IsEmpty(Cells(i, "A")) Then
                        If Cells(i, "BB").Interior.Color = RGB(160, 160, 530) Then
                            MsgBox "Calculation already preformed"
                                Exit Sub
                        End If
                        If IsNumeric(Cells(i, "BB")) Then
                            If Cells(i, "E").Value2 = "Dissolver" Then
                                Cells(i, "BB") = Cells(i, "BB") * Cells(i, "F")
                                Cells(i, "BB").Interior.Color = RGB(0, 150, 520)
                            Else: Cells(i, "BB") = (Cells(i, "BB") * Cells(i, "F") * Cells(i, "G"))
                                Cells(i, "BB").Interior.Color = RGB(160, 160, 530)
                            End If
                        End If
                    
                    End If
                    
        i = i + 1
    
    Loop
   
   
        
        'Determines the range in which the data lies
        Lrow = Range("H" & Rows.Count).End(xlUp).Row
        Lcol = Cells(27, Columns.Count).End(xlToLeft).Column
    
            Set RNG = Range(Cells(27, "F"), Cells(Lrow, Lcol))
                
                'Sets the output to scientific notation and changes font size and style
                With RNG
                    .NumberFormat = "0.00E+00"
                    .Font.Name = "Times New Roman"
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                End With
                
                For Each Cell In RNG
                    If Application.WorksheetFunction.IsNA(Cell) Then
                        Cell.Value2 = ""
                        Cell.NumberFormat = "0"
                    End If
                Next

                'Automaticaly sizes the columns to fit
                With Sheets("Sample Totals")
                    .Columns.AutoFit
                End With
Exit Sub

Error_Handler:
    MsgBox "Missing Data. Ensure all relavent data is entered", vbExclamation
    
End Sub
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 29 March 2018
    'Calculates a dilution factor from data in dilutin sheets and dissolver wt.

Sub Dilution()
Dim Lrow As Long
Dim Frow As Long
Dim RNG As Range
Dim A As Long
Dim B As Long
Dim Answer As Integer

'Selects the appropriate sheet to run Macro
Sheets("Dilutions").Select

'Operator must acknoldge all data has been enterd
Answer = MsgBox("Has the type of dilution been selected in column E (Initial or Serial)? Have all dilution values been entered correctly?", vbYesNo + vbQuestion, "Input Data")
    If Answer = vbNo Then
        End
        Else
    End If

    'Sets First and Last row of data to be anlayzed
    Frow = 3
    Lrow = Range("A" & Rows.Count).End(xlUp).Row
        Cells(Lrow, "A").Interior.Color = RGB(340, 200, 240)
    
    'Steps the data from bottom to top
    For Lrow = Lrow To Frow Step -1
           
        'Looks in row E for the number of dilution performed to get to the one being disspossed
        With Cells(Lrow, "E")
        
        'IF error occurs due to missing data Sub will Exit and display message
        On Error GoTo Error_Handler:
            
            'Calculation is perfomed based on description provided by column E
            If .Value2 = "Dissolver" Then
                Cells(Lrow, "F") = 1
            End If
            
            If .Value2 = "Initial" Then
                Cells(Lrow, "Z") = Cells(Lrow, "H") / (Cells(Lrow, "I") * Cells(Lrow, "D"))
                .Interior.Color = RGB(140, 150, 240)
                
            End If
            
            If .Value2 = "Serial" Then
                Cells(Lrow, "AA") = (Cells(Lrow, "H") * Cells(Lrow, "J")) / (Cells(Lrow, "I") * Cells(Lrow, "K") * Cells(Lrow, "D"))
                .Interior.Color = RGB(140, 150, 240)
                
            End If
                      
        End With
        
        If Cells(2, "I").Value2 = "Dilution Mass (g)" And Cells(Lrow, "E").Value2 = "Initial" Then
            Cells(Lrow, "F") = Cells(Lrow, "Z").Value2 * Cells(Lrow, "G").Value2
            Cells(Lrow, "F").Interior.Color = RGB(240, 240, 100)
        ElseIf Cells(2, "I").Value2 = "Dilution Volume (mL)" And Cells(Lrow, "E").Value2 = "Initial" Then
            Cells(Lrow, "F") = Cells(Lrow, "Z").Value
            Cells(Lrow, "F").Interior.Color = RGB(140, 150, 240)
        End If
        
        If Cells(2, "I").Value2 = "Dilution Mass (g)" And Cells(Lrow, "E").Value2 = "Serial" Then
            Cells(Lrow, "F") = Cells(Lrow, "AA").Value2 * Cells(Lrow, "G").Value2
            Cells(Lrow, "F").Interior.Color = RGB(240, 200, 240)
        ElseIf Cells(2, "I").Value2 = "Dilution Volume (mL)" And Cells(Lrow, "E").Value2 = "Serial" Then
            Cells(Lrow, "F") = Cells(Lrow, "AA").Value
            Cells(Lrow, "F").Interior.Color = RGB(140, 150, 240)
        End If
        
    Next Lrow
            
    Sheets("Dilutions").Select
            'Sets First and Last row of data to be anlayzed
    Frow = 3
    Lrow = Range("A" & Rows.Count).End(xlUp).Row
        Cells(Lrow, "A").Interior.Color = RGB(340, 200, 240)
    
    'Steps the data from bottom to top
    For Lrow = Lrow To Frow Step -1
    
        If Cells(Lrow, "F").Value2 = 0 Then
            MsgBox "Incorrect calculation perfomed. Ensure all aliquot Mass/Volumes are entered (on the Dilutions Tab) for dilution, and reinitiate Macro."
        End If
        
    Next Lrow
    
    'Determines the range in which the data lies
    Lrow = Range("I" & Rows.Count).End(xlUp).Row
        Set RNG = Range(Cells(3, "F"), Cells(Lrow, "F"))
                
            'Sets the output to scientific notation and changes font size and style
            With RNG
                .NumberFormat = "0.00E+00"
                .Font.Name = "Times New Roman"
                .Font.Size = 11
                .HorizontalAlignment = xlCenter
            End With

            'Automaticaly sizes the columns to fit
            With Sheets("Sample Totals")
                .Columns.AutoFit
            End With
            
    Columns("Z:AA").Delete
   
Exit Sub

Error_Handler:
    MsgBox "Missing Data on Dilutions Sheet. Ensure all relavent data is entered"
    End
    
End Sub
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 02 April 2018
    'Moves the Dissolver wt. to the Sample Totals Tab
    
Sub Transfer_Dwt()
Dim i As Integer
Dim Lrow As Long
Dim Frow As Long

i = 2

Sheets("Raw Data").Select

            Columns("E:E").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("E1").Select
                    ActiveCell.FormulaR1C1 = "HelperTab"
                    
        Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))

                Cells(i, "E").FormulaR1C1 = "=RC[-4]&RC[-3]&RC[-1]"
                
        i = i + 1
        
        Loop
        
Sheets("Sample Totals").Select

    'Sets First and Last row of data to be anlayzed
    Frow = 27
    Lrow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Steps the data from bottom to top
    For Lrow = Lrow To Frow Step -1
    
        Cells(Lrow, "C") = "=VLOOKUP('Sample Totals'!RC1&'Sample Totals'!RC2&'Sample Totals'!R26C3,'Raw Data'!C5:C6,2,FALSE)"
        
    Next Lrow
    
i = 27

    Sheets("Sample Totals").Select
    
        'Determines the range in which the data lies
        Lrow = Range("C" & Rows.Count).End(xlUp).Row
        Lcol = Cells(27, Columns.Count).End(xlToLeft).Column
    
            Set RNG = Range(Cells(27, "C"), Cells(Lrow, "C"))
              
                'Sets the output to scientific notation and changes font size and style
                With RNG
                    .Copy
                End With
                Cells(27, "C").PasteSpecial xlPasteValues
                    
Sheets("Raw Data").Select
Columns("E:E").Delete
Sheets("Sample Totals").Select

End Sub
