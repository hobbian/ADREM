Attribute VB_Name = "Remove_Blanks"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 15 May 2015
    'Subroutine removes rows with blank in the Results section

Sub Delete_Blanks()

Dim Lrow As Long
Dim Firstrow As Long
Dim Lastrow As Long
Dim ViewMode As Long
Dim CalcMode As Long

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

        'Delete Non Empricial Data From Row E
        'Looks for data in worksheet labled Raw Data
        With Sheets("Raw Data")
        'Selects the sheet
            .Select
        'Changes the view to normal view
            ViewMode = ActiveWindow.View
            ActiveWindow.View = xlNormalView
        'Turn off Page Braks for Speed
            .DisplayPageBreaks = False
            'Set the First and Last row to loop through
            Firstrow = 2
            Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        'We loop from Lastrow to Firstrow
            For Lrow = Lastrow To Firstrow Step -1
                
                'Chcek the values in the E column
                With .Cells(Lrow, "E")
                            
                        If .Value = "Cancel" Or .Value = "n/a" Or .Value = "N/A" Or .Value = "NA" Or .Value = "ND" Or .Value = "cancel" Or .Value = "CANCEL" Or .Value = "" Then .EntireRow.Delete
                   
                 End With
                 
            Next Lrow
                
        'Set the First and Last row to loop through
        Firstrow = 2
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        'We loop from Lastrow to Firstrow
            For Lrow = Lastrow To Firstrow Step -1
                
                'Chcek the values in the E column
                With .Cells(Lrow, "E")
                            
                        If .Value = "Cancel" Or .Value = "n/a" Or .Value = "N/A" Or .Value = "NA" Or .Value = "ND" Or .Value = "cancel" Or .Value = "CANCEL" Or .Value = "" Then .EntireRow.Delete
                        If .Value = "Saturated" Then .EntireRow.Delete
                        
                 End With
                 
            Next Lrow
    
    End With
            
End Sub
