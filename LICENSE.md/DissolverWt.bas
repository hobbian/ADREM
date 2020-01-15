Attribute VB_Name = "DissolverWt"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 23 May 2018
    'Moves the Dissolver wt. to the Dilutions Tab
    
Sub Transfer_Dwt_Dilution()
Dim i As Integer
Dim Lrow As Long
Dim Frow As Long
Dim Lcol As Long
Dim RNG As Range

    
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
        
Sheets("Dilutions").Select

    'Sets First and Last row of data to be anlayzed
    Frow = 3
    Lrow = Range("A" & Rows.Count).End(xlUp).Row
    
    'Steps the data from bottom to top
    For Lrow = Lrow To Frow Step -1
    
        Cells(Lrow, "D") = "=VLOOKUP('Dilutions'!RC1&'Dilutions'!RC2&'Dilutions'!R2C,'Raw Data'!C5:C6,2,FALSE)"
        Cells(Lrow, "C") = "=VLOOKUP('Dilutions'!RC1&'Dilutions'!RC2&'Dilutions'!R2C,'Raw Data'!C5:C6,2,FALSE)"
            If IsError(Cells(Lrow, "C")) Then Cells(Lrow, "C") = "=VLOOKUP('Dilutions'!RC1&'Dilutions'!RC2&'Dilutions'!R1C,'Raw Data'!C5:C6,2,FALSE)"
            If IsError(Cells(Lrow, "C")) Then Cells(Lrow, "C").ClearContents
         
    Next Lrow
    
    Sheets("Dilutions").Select
    
        'Determines the range in which the data lies
        Lrow = Range("C" & Rows.Count).End(xlUp).Row
        Lcol = Cells(3, "D")
    
            Set RNG = Range(Cells(3, "C"), Cells(Lrow, Lcol))
              
                'Sets the output to scientific notation and changes font size and style
                With RNG
                    .Copy
                End With
                    Cells(3, "C").PasteSpecial xlPasteValues
                    
    Sheets("Raw Data").Select
    Columns("E:E").Delete
    Sheets("Dilutions").Select


            With Columns("C:D")
                .NumberFormat = "0.0000"
                .Font.Name = "Times New Roman"
                .Font.Size = 11
                .HorizontalAlignment = xlCenter
            End With
            
            With Columns("G:J")
                .NumberFormat = "0.0000"
                .Font.Name = "Times New Roman"
                .Font.Size = 11
                .HorizontalAlignment = xlCenter
            End With


End Sub
