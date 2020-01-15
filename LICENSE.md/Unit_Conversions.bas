Attribute VB_Name = "Unit_Conversions"
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 08 May 2018
    'Used to convert Wt% for U and Pu to a weight in grams for consolidations.

Sub WTPercent_Conversion()
Dim i As Integer
Dim RNG As Range
Dim RNG_1 As Range
Dim RNG_2 As Range
Dim ALNumber As Long
Dim Sample_No As Long
Dim PERCENT As Single
Dim GetRow As Long
Dim CalcMode As Long
Dim ReplaceSPL As String
Dim Line As String
Dim ViewMode As Long
Dim Error As Integer

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = True
    End With
        With Sheets("Raw Data")
        'Selects the sheet
            .Select
        'Changes the view to normal view
            ViewMode = ActiveWindow.View
            ActiveWindow.View = xlNormalView
        'Turn off Page Braks for Speed
            .DisplayPageBreaks = True
        End With
        
i = 2
         
On Error GoTo ErrorHandlerWtP:

    'Define Loop and set variables for the data rows to convert
    Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
            
    
        With Sheets("Raw Data")
        ALNumber = Cells(i, "A").Value
        PERCENT = Cells(i, "E").Value / 100
        Sample_No = Cells(i, "B").Value
        
            'Use the autofilter to find the Pu total wt
            If Cells(i, "F").Value = "Wt%" And Cells(i, "D") = "Pu Total" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Physical Measurements"
                    .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt."
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                   
                'If a wt./wt. concentration is detected with no sample weight it is assumed it is g analyte per g diluent
                    'This portion converts the dilution Wt. to a Spl. Wt.
                    If RNG.Value = "RESULT" Then
                        .Range("A1:D1").AutoFilter Field:=4
                            ReplaceSPL = Application.InputBox("Could not convert concentration to total grmas for the AL Number. Select the appropriate Species from column D to replace Spl. Wt. for the given AL Number.", Title:="Spl. Wt. Replacement", Type:=8)
                            
                            'If user selects cancel the subroutine will exit
                            If ReplaceSPL = "False" Then
                                MsgBox "Macro incomplete. Input needed data and reinitiate", , "Macro Incomplete"
                                End
                            End If
                            
                        Set RNG_1 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=ReplaceSPL, Replacement:="Spl. Wt.", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                                                        
                        'Applys the filter once the Spl. Wt. has been applied
                        Set RNG = Nothing
                            .AutoFilterMode = False
                            .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                            .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                            .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Physical Measurements"
                            .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt."
                        Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Interior.Color = RGB(115, 50, 160)
                            RNG_2.Font.Color = RGB(255, 255, 255)
                        
                    End If
                
                'Calculate the weight and replace it and the ISO% with g.
                Cells(i, "E").Formula = "=" & Cells(i, "E").Value / 100 * RNG.Value
                    Cells(i, "F").Formula = "g"
                    .AutoFilterMode = False
                    Cells(i, "E").Interior.Color = RGB(159, 248, 110)
                    Cells(i, "F").Interior.Color = RGB(159, 248, 110)
            End If
        
        'Use the autofilter to find the U total wt
            If Cells(i, "F").Value = "Wt%" And Cells(i, "D") = "U Total" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Physical Measurements"
                    .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt."
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    
                'If a wt./wt. concentration is detected with no sample weight it is assumed it is g analyte per g diluent
                    'This portion converts the dilution Wt. to a Spl. Wt.
                    If RNG.Value = "RESULT" Then
                        .Range("A1:D1").AutoFilter Field:=4
                            ReplaceSPL = Application.InputBox("Could not convert concentration to total grmas for the AL Number. Select the appropriate Species from column D to replace Spl. Wt. for the given AL Number.", Title:="Spl. Wt. Replacement", Type:=8)
                        
                            'If user selects cancel the subroutine will exit
                            If ReplaceSPL = "False" Then
                                MsgBox "Macro incomplete. Input needed data and reinitiate", , "Macro Incomplete"
                                End
                            End If
                            
                        Set RNG_1 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=ReplaceSPL, Replacement:="Spl. Wt.", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                                                        
                        'Applys the filter once the Spl. Wt. has been applied
                        Set RNG = Nothing
                            .AutoFilterMode = False
                            .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                            .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Physical Measurements"
                            .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt."
                        Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Interior.Color = RGB(115, 50, 160)
                            RNG_2.Font.Color = RGB(255, 255, 255)
                        
                    End If
                
                'Calculate the weight and replace it and the ISO% with g.
                Cells(i, "E").Formula = "=" & Cells(i, "E").Value / 100 * RNG.Value
                    Cells(i, "F").Formula = "g"
                    .AutoFilterMode = False
                    Cells(i, "E").Interior.Color = RGB(159, 248, 110)
                    Cells(i, "F").Interior.Color = RGB(159, 248, 110)
            End If
        
        End With
        
      i = i + 1
      
    Loop

i = 2

    'Define Loop and set variables for the data rows to convert
    Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
            
    
        With Sheets("Raw Data")
            ALNumber = Cells(i, "A").Value
            PERCENT = Cells(i, "E").Value / 100
            Sample_No = Cells(i, "B").Value
        
                'Use the autofilter to find the Pu total wt
                If Cells(i, "F").Value = "Wt%" And Right(Cells(i, "D"), 2) = "Pu" Then
                    Set RNG = Nothing
                        .AutoFilterMode = False
                        .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                        .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                        .Range("A1:D1").AutoFilter Field:=3, Criteria1:=Cells(i, "C")
                        .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Pu Total"
                    Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                        
                    
                    'Calculate the weight and replace it and the ISO% with g.
                    Cells(i, "E").Formula = "=" & Cells(i, "E").Value / 100 * RNG.Value
                        Cells(i, "F").Formula = "g"
                        .AutoFilterMode = False
                        Cells(i, "E").Interior.Color = RGB(159, 248, 110)
                        Cells(i, "F").Interior.Color = RGB(159, 248, 110)
                End If
            
            'Use the autofilter to find the U total wt
                If Cells(i, "F").Value = "Wt%" And Right(Cells(i, "D"), 1) = "U" Then
                    Set RNG = Nothing
                        .AutoFilterMode = False
                        .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                        .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                        .Range("A1:D1").AutoFilter Field:=3, Criteria1:=Cells(i, "C")
                        .Range("A1:D1").AutoFilter Field:=4, Criteria1:="U Total"
                    Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                        
                    
                    'Calculate the weight and replace it and the ISO% with g.
                    Cells(i, "E").Formula = "=" & Cells(i, "E").Value / 100 * RNG.Value
                        Cells(i, "F").Formula = "g"
                        .AutoFilterMode = False
                        Cells(i, "E").Interior.Color = RGB(159, 248, 110)
                        Cells(i, "F").Interior.Color = RGB(159, 248, 110)
                End If
        
        End With
        
      i = i + 1
      
    Loop

Exit Sub

ErrorHandlerWtP:
    Erorr = MsgBox("Does the data correspond to the headers.", vbYesNo + vbQuestion, "Error Message")
        If Error = vbYes Then
            MsgBox "Error Exicuting the macro. Fix error and reinitiate.", vbExclimation, "Error"
            End
        End If
        
        If Error = vbNo Then End
        
End Sub
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Joey Charboneau February 27, 2018
'Ian M. Hobbs March 5, 2018
    'Minor Modification to only look at rows with associated AL numbers
    'Used to convert ISO% for U and Pu to a weight in grams for consolidations.
    
Sub ISO_Conversion()
Dim i As Integer
Dim RNG As Range
Dim ALNumber As Long
Dim PERCENT As Single
Dim GetRow As Long
Dim CalcMode As Long
Dim Sample_No As Integer
i = 2
       
'Define Loop and set variables for the data rows to convert
Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
        

    With Sheets("Raw Data")
    ALNumber = Cells(i, "A").Value
    PERCENT = Cells(i, "E").Value / 100
    Sample_No = Cells(i, "B").Value
    
        'Use the autofilter to find the Pu total wt
        If Cells(i, "F").Value = "ISO%" And Right(Cells(i, "D"), 2) = "Pu" Then
            Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:D1").AutoFilter Field:=3, Criteria1:=Cells(i, "C")
                .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Pu Total"
            Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
            'Calculate the weight and replace it and the ISO% with g.
            Cells(i, "E").Formula = "=" & Cells(i, "E").Value / 100 * RNG.Value
                Cells(i, "F").Formula = "g"
                .AutoFilterMode = False
                Cells(i, "E").Interior.Color = RGB(159, 248, 110)
                Cells(i, "F").Interior.Color = RGB(159, 248, 110)
        End If
    
    'Use the autofilter to find the U total wt
        If Cells(i, "F").Value = "ISO%" And Right(Cells(i, "D"), 1) = "U" Then
            Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:D1").AutoFilter Field:=3, Criteria1:=Cells(i, "C")
                .Range("A1:D1").AutoFilter Field:=4, Criteria1:="U Total"
            Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
            
            'Calculate the weight and replace it and the ISO% with g.
            Cells(i, "E").Formula = "=" & Cells(i, "E").Value / 100 * RNG.Value
                Cells(i, "F").Formula = "g"
                .AutoFilterMode = False
                Cells(i, "E").Interior.Color = RGB(159, 248, 110)
                Cells(i, "F").Interior.Color = RGB(159, 248, 110)
        End If
    
    End With
    
  i = i + 1
  
Loop

End Sub
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 08 May 2018
    'Used to convert µCi/g concentations to total Ci.

Sub ActivityConc_Conversion()
Dim i As Integer
Dim RNG As Range
Dim RNG_1 As Range
Dim RNG_2 As Range
Dim ALNumber As Long
Dim Sample_No As Long
Dim PERCENT As Single
Dim GetRow As Long
Dim CalcMode As Long
Dim ReplaceSPL As String
Dim Line As String


i = 2
  
   
       'Define Loop and set variables for the data rows to convert
Ln155: Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
             
    
        With Sheets("Raw Data")
        ALNumber = Cells(i, "A").Value
        PERCENT = Cells(i, "E").Value / 100
        Sample_No = Cells(i, "B").Value
        
            'Use the autofilter to find the Pu total wt
            If Cells(i, "F").Value = "µCi/g" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Physical Measurements"
                    .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt.", Operator:=xlAnd
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                                               
                    'If a wt./wt. concentration is detected with no sample weight it is assumed it is g analyte per g diluent
                    'This portion converts the dilution Wt. to a Spl. Wt.
                    If RNG.Value = "RESULT" Then
                        .Range("A1:D1").AutoFilter Field:=4
                            ReplaceSPL = Application.InputBox("Could not convert concentration to total grmas for the AL#" & ALNumber & ". Check to ensure no Spl. Wt. exists but was not recorded in LIMS. Select the appropriate Species from column D to replace Spl. Wt. for the given AL Number.", Title:="Spl. Wt. Replacement", Type:=8)
                                
                                'If user selects cancel the subroutine will exit
                                If ReplaceSPL = "False" Then
                                    MsgBox "Macro incomplete. Input needed data and reinitiate", , "Macro Incomplete"
                                    End
                                End If
                                    
                        Set RNG_1 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=ReplaceSPL, Replacement:="Spl. Wt.", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                                                        
                        'Applys the filter once the Spl. Wt. has been applied
                        Set RNG = Nothing
                            .AutoFilterMode = False
                            .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                            .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                            .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt."
                        Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Interior.Color = RGB(115, 50, 160)
                            RNG_2.Font.Color = RGB(255, 255, 255)
                        
                    End If
                                                       
            'Calculate the weight and replace it and the ISO% with g.
                Cells(i, "E").Formula = "=" & Cells(i, "E").Value * RNG.Value / 1000000
                Cells(i, "F").Formula = "Ci"
                .AutoFilterMode = False
                Cells(i, "E").Interior.Color = RGB(159, 248, 110)
                Cells(i, "F").Interior.Color = RGB(159, 248, 110)
            End If
           
        End With
        
      i = i + 1
      
    Loop
Exit Sub

'If multiple weight occur for a single sample set one must be removed for macro to proceed
ErrrorHandler_WeightConc_Conversion:
    
    MsgBox "Species not selected, reinitiate macro and select the appropriate macro"

End Sub
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 07 May 2018
    'Used to convert wt./wt. concentations to totals.

Sub WeightConc_Conversion()
Dim i As Integer
Dim RNG As Range
Dim RNG_1 As Range
Dim RNG_2 As Range
Dim ALNumber As Long
Dim Sample_No As Long
Dim PERCENT As Single
Dim GetRow As Long
Dim CalcMode As Long
Dim ReplaceSPL As String
Dim Line As String
Dim C As Range

i = 2
  
   
'Define Loop and set variables for the data rows to convert
Ln155: Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
         

    With Sheets("Raw Data")
    ALNumber = Cells(i, "A").Value
    PERCENT = Cells(i, "E").Value / 100
    Sample_No = Cells(i, "B").Value
    
        'Use the autofilter to find the Pu total wt
        If Cells(i, "F").Value = "µg/g" Then
            Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:D1").AutoFilter Field:=3, Criteria1:="Physical Measurements"
                .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt.", Operator:=xlAnd
            Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                                           
                'If a wt./wt. concentration is detected with no sample weight it is assumed it is g analyte per g diluent
                'This portion converts the dilution Wt. to a Spl. Wt.
                If RNG.Value = "RESULT" Then
                    .Range("A1:D1").AutoFilter Field:=4
                        ReplaceSPL = Application.InputBox("Could not convert concentration to total grmas for the AL#" & ALNumber & ". Select the appropriate Species from column D to replace Spl. Wt. for the given AL Number.", Title:="Spl. Wt. Replacement", Type:=8)
                    
                            'If user selects cancel the subroutine will exit
                            If ReplaceSPL = "False" Then
                                MsgBox "Macro incomplete. Input needed data and reinitiate", , "Macro Incomplete"
                                End
                            End If
                                
                    Set RNG_1 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                        RNG_1.Replace What:=ReplaceSPL, Replacement:="Spl. Wt.", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                                                    
                    'Applys the filter once the Spl. Wt. has been applied
                    Set RNG = Nothing
                        .AutoFilterMode = False
                        .Range("A1:D1").AutoFilter Field:=1, Criteria1:=ALNumber
                        .Range("A1:D1").AutoFilter Field:=2, Criteria1:=Sample_No
                        .Range("A1:D1").AutoFilter Field:=4, Criteria1:="Spl. Wt."
                    Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                        RNG_2.Interior.Color = RGB(115, 50, 160)
                        RNG_2.Font.Color = RGB(255, 255, 255)
                    
                End If
                                                   
        'Calculate the weight and replace it and the ISO% with g.
            Cells(i, "E").Formula = "=" & Cells(i, "E").Value * RNG.Value / 1000000
            Cells(i, "F").Formula = "g"
            .AutoFilterMode = False
            Cells(i, "E").Interior.Color = RGB(159, 248, 110)
            Cells(i, "F").Interior.Color = RGB(159, 248, 110)
        End If
       
    End With
    
  i = i + 1
  
Loop
Exit Sub

'If multiple weight occur for a single sample set one must be removed for macro to proceed
ErrrorHandler_WeightConc_Conversion:
    
    MsgBox "Species not selected, reinitiate macro and select the appropriate macro"

End Sub

