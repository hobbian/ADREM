Attribute VB_Name = "Less_Than_Removal_Sub"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine Written by Ian M. Hobbs 18 April 2018
    'Subroutine removes the values that where previously less thans
    'RCRA metal less thans are retained a maximium values

Sub Less_Than_Removal()
Dim Lrow As Integer
Dim Deleted As Boolean
Dim Firstrow As Integer
Dim Lastrow As Long
Dim CalcMode As Long
Dim ViewMode As Long

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    With Sheets("Raw Data")
        'Selects the sheet
            .Select
        'Changes the view to normal view
            ViewMode = ActiveWindow.View
            ActiveWindow.View = xlNormalView
        'Turn off Page Braks for Speed
            .DisplayPageBreaks = False
    End With
        
    'Selects the approptiate sheet
    With Sheets("Raw Data")
        
        'Set the First and Last row to loop through
        Firstrow = 2
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
            
            'Selects the isotopes relevant for tracking
            'Modified data will not be deleted due to replicate ElseIf statements
            For Lrow = Lastrow To Firstrow Step -1
           
                If Cells(Lrow, "D") = "54Mn" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "60Co" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "90Sr" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "90Y" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "90M/Z" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "99Tc" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "106Ru/Rh" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "134Cs" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "137Cs" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "144Ce" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "154Eu" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "155Eu" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "233U" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "234U" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "235U" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "236U" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "237Np" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "238U" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "U Total" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "238Pu" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "238M/Z" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "239Pu" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "240Pu" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "241Pu" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "242Pu" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "241Am" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "Pu Total" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "241M/Z" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "243Am" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                If Cells(Lrow, "D") = "244Cm" And Cells(Lrow, "I") = "Value was Converted from a <Value" Then Cells(Lrow, "D").EntireRow.Delete
                             
            Next Lrow

    End With
    
    'Makes it pretty
    With Sheets("Raw Data")
        .Columns.AutoFit
    End With
  
End Sub

