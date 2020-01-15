Attribute VB_Name = "Dilution_Transfers"
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine Written by Ian M. Hobbs 02 April 2018
'Subroutine transfers the dilution factors from the
    'Dilution Tab to the Sample Totals Tab

Sub Transfer_DF()
Dim i As Integer
Dim j As Integer
Dim Lrow As Integer
Dim Frow As Integer
Dim RNG As Range

i = 3
j = 3

'Generates a helper tab for Vlookup function
Sheets("Dilutions").Select
            Columns("F:F").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("F1").Select
                    ActiveCell.FormulaR1C1 = "HelperTab"
                    
        Do While Not IsEmpty(Sheets("Dilutions").Cells(i, "A"))

                Cells(i, "F").FormulaR1C1 = "=RC[-5]&RC[-4]&RC[-1]"
                
        i = i + 1
        
        Loop


i = 27

    'Uses the Helper Tab to Transfer dilution factor to
    'Sample Totals worksheet
    Sheets("Sample Totals").Select
    
        'Stops the function once the data set is complete
        Do While Not IsEmpty(Sheets("Sample Totals").Cells(i, "A"))
              
              'Skips Solution that are dissolvers
              If Not Cells(i, "E") = "Dissolver" Then
                
                
                If Cells(i, "F") >= 0.002 Then
              
                    'Inputs data corresponding to the AL# the Sample ID and the Element/Isotope
                    'and paste it on Sample Totals Tab
                    Lookup_1 = Sheets("Sample Totals").Cells(i, "A").Value2
                    Lookup_2 = Sheets("Sample Totals").Cells(i, "B").Value2
                    Lookup_3 = Sheets("Sample Totals").Cells(i, "E").Value2
                    Lookup_4 = Lookup_1 & Lookup_2 & Lookup_3
                    Sheets("Sample Totals").Cells(i, "F") = Application.VLookup(Lookup_4, Sheets("Dilutions").Range("F3:G100"), 2, False)
                
                End If
                
            End If
            
            If Cells(i, "E") = "Dissolver" Then Cells(i, "F") = 1
            
        i = i + 1
        
        Loop

i = 27
    
    'Checks data to ensure dilution factor was accuratly copied
    Do While Not IsEmpty(Sheets("Sample Totals").Cells(i, "A"))
        
        On Error GoTo Error_Handler:
        
        If Cells(i, "F") = 0 Then
             MsgBox "Data missing from Dilutions Sheet! :( Enter data and rerun macro."
       End If
            
        i = i + 1
        
    Loop


'Copies Vlookup data and paste values so when Helper Tab is
'Deleted the values will not change to N/A
Sheets("Sample Totals").Select

        'Determines the range in which the data lies
        Lrow = Range("F" & Rows.Count).End(xlUp).Row
        
            'Selects range of the values added
            Set RNG = Range(Cells(27, "F"), Cells(Lrow, "F"))
                
                'Copies function and pastes the values
                With RNG
                    .Copy
                End With
                Cells(27, "F").PasteSpecial xlPasteValues
                
                'Sets the output to scientific notation and changes font size and style
                With RNG
                    .NumberFormat = "0.00E+00"
                    .Font.Name = "Times New Roman"
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                End With
                

    Sheets("Dilutions").Select
    Columns("F:F").Delete
Cells(2, "F") = "Dilution Factor"
    Sheets("Sample Totals").Select
    
Exit Sub

Error_Handler:
    MsgBox "Data mismatch between Sample Totals and Dilutions worksheet insure AL# Sample ID and Type of Dillutions for all data on Sample Totals is represent on Dilutions worksheet and vis versa. :O", vbExclimation
    Sheets("Dilutions").Select
    Columns("F:F").Delete
    Sheets("Sample Totals").Select
    End

End Sub

'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 02 April 2018
'Modidifies the formating on the Sample Totals Tab

Sub Format()
Dim RNG1 As Range
Dim RNG2 As Range
Dim RNG3 As Range
Dim RNG4 As Range

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = False
    End With
            
    'Copies Vlookup data and paste values so when Helper Tab is
    'Deleted the values will not change to N/A
    Sheets("Sample Totals").Select

        'Determines the range in which the data lies
        Lrow = Range("F" & Rows.Count).End(xlUp).Row
        
            Set RNG1 = Range(Cells(27, "A"), Cells(Lrow, "A"))
                
                'Sets the output to scientific notation and changes font size and style
                With RNG1
                    .Font.Name = "Times New Roman"
                    .Font.Size = 11
                    .Font.Bold = True
                    .HorizontalAlignment = xlLeft
                End With
                
            
            Set RNG2 = Range(Cells(27, "C"), Cells(Lrow, "D"))
                
                'Sets the output to scientific notation and changes font size and style
                With RNG2
                    .NumberFormat = "0.0000"
                    .Font.Name = "Times New Roman"
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                End With
                
            Set RNG3 = Range(Cells(27, "B"), Cells(Lrow, "B"))
                
                'Sets the output to scientific notation and changes font size and style
                With RNG3
                    .NumberFormat = "0"
                    .Font.Name = "Times New Roman"
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                End With
                
            Set RNG4 = Range(Cells(27, "E"), Cells(Lrow, "E"))
                
                'Sets the output to scientific notation and changes font size and style
                With RNG4
                    .NumberFormat = "0"
                    .Font.Name = "Times New Roman"
                    .Font.Size = 11
                    .HorizontalAlignment = xlCenter
                End With
        
End Sub
