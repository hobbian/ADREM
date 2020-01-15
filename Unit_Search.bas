Attribute VB_Name = "Unit_Search"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 17 May 2018
    'Performs a final check to insure all units are in g

Sub Unit_Confirm()
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
                
                'Chcek the values in the F column for gross alpha and deletes
                With .Cells(Lrow, "F")
                            
                        If Not .Value = "g" Then
                            MsgBox "Invalid Unit Detected!. Filter data to find and correct unit", vbExclamation, "Unit Error"
                            End
                        End If
                                           
                 End With
                 
            Next Lrow
    
    End With
            
End Sub
