Attribute VB_Name = "Duplicate_Removal"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 14 May 2018
    'Removes the duplicate values based on the defined

Sub Remove_Duplicates()
Dim ViewMode As Long
Dim CalcMode As Long

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    

    'Looks for data in worksheet labled Raw Data
    With Sheets("Raw Data")
    'Selects the sheet
        .Select
    'Changes the view to normal view
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView
    'Turn off Page Braks for Speed
        .DisplayPageBreaks = False
    
        'Removes duplicate data based on ALNumber, SampleNo, Measure, Species, Result, Units, and Date completed
        Sheets("Raw Data").Select
            Range("A1").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
            ActiveSheet.Range("$A$1:$H$46900").RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8), Header:=xlYes
            
    End With
    
End Sub
