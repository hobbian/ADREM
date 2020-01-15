Attribute VB_Name = "Remove_Less_Conservative_Value"
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine was writen by Ian M. Hobbs 05 May 2018
    'This subroutine is used to delete the less conservative value from the dataset

Sub Remove_Non_Conservative_Value()
Dim i As Integer
Dim RNG As Range
Dim ALNumber As Long
Dim ISOTOPE As Long
Dim MINTOMAX As Long
Dim CalcMode As Long
Dim ViewMode As Long
Dim Sample_No As Integer
Dim Answer As Integer

Answer = MsgBox("This subroutine will remove the smaller value for a element/isotope if mulitiple values are reported. Do you wish to proceed?", vbYesNo + vbQuestion, "Keeping the Conservative Value for the Sample")
    If Answer = vbNo Then End
    
i = 2

    'Turn off Sceen Updating Make (single screen update at the end of the caluclation)
        With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
           
            
            'Changes the view to normal view
            ViewMode = ActiveWindow.View
            ActiveWindow.View = xlNormalView


'Takes Data Miner to the approptiate worksheet
Worksheets("Raw Data").Activate

    'Selects the sheet to perform the macro on
    With Sheets("Raw Data")
    
       'Selects lines that have an associated AL number
        Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
        
        'Defines the columns autofilter will look through
        ALNumber = Cells(i, "A").Value2
        Sample_No = Cells(i, "B").Value2
    
            'Use the autofilter function to find the most conservative measurement perfomred on the isotope in question
            If Cells(i, "D") = "240Pu" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="240Pu"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "234U" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="234U"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                         Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "235U" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="235U"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "236U" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="236U"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "237Np" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="237Np"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "238U" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="238U"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "238Pu" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="238Pu"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "239Pu" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="239Pu"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "241Pu" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="241Pu"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "242Pu" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="242Pu"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "241Am" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="241Am"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If

            If Cells(i, "D") = "243Am" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="241Am"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If

            If Cells(i, "D") = "244Cm" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="241Am"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "99Tc" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="99Tc"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Ag" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "As" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Ba" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            

            
            If Cells(i, "D") = "Be" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Cd" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Cd"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Cr" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Cr"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Hg" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Ni" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Pb" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Se" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Sb" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "Tl" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="Ba"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "54Mn" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="54Mn"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "60Co" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="60Co"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "90Sr" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="90Sr"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
                       
            If Cells(i, "D") = "90Y" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="90Y"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                        
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
                        
            If Cells(i, "D") = "90M/Z" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="90Y"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
                      
            If Cells(i, "D") = "106Ru/Rh" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="106Ru/Rh"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
                        
            If Cells(i, "D") = "125Sb" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="125Sb"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "134Cs" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="134Cs"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
                        If Cells(i, "D") = "137Cs" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="137Cs"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
                        If Cells(i, "D") = "144Ce" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="144Ce"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
                If Cells(i, "D") = "154Eu" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="154Eu"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                
                'Turns off the autofilter
                .AutoFilterMode = False
            
            End If
            
            If Cells(i, "D") = "155Eu" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="155Eu"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                              
            
            'Turns off the autofilter
                .AutoFilterMode = False
                
            End If
            
            If Cells(i, "D") = "243Am" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="243Am"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                                    
                'Turns off the autofilter
                .AutoFilterMode = False
                
            End If
                        
            If Cells(i, "D") = "244Cm" Then
                Set RNG = Nothing
                .AutoFilterMode = False
                .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                .Range("A1:H1").AutoFilter Field:=4, Criteria1:="244Cm"
                Set RNG = Range("E1:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                
                    'Determins the Max and Min values for the Range described by RNG
                    Min1 = Application.WorksheetFunction.Min(RNG)
                    Max1 = Application.WorksheetFunction.Max(RNG)
                  
                        'Describes Minimum value
                        Cells(i, "L") = Min1
                            If Cells(i, "E") = Min1 Then
                            Cells(i, "K") = "Min"
                        
                        End If
                  
                        'Describes the Maximum value
                        Cells(i, "M") = Max1
                            If Cells(i, "E") = Max1 Then
                            Cells(i, "K") = "Max"
                            Cells(i, "E").Interior.Color = RGB(200, 200, 0)
                            
                        End If
                                
                'Turns off the autofilter
                .AutoFilterMode = False
                
            End If
          
          i = i + 1
          
       Loop
       
           
    'Set the First and Last row to loop through
            Firstrow = 2
            Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
        'We loop from Lastrow to Firstrow
            For Lrow = Lastrow To Firstrow Step -1
                
                           
                'Chcek the values in the E column
                With .Cells(Lrow, "K")
                            
                        If .Value Like "Min" Then .EntireRow.Delete
                   
                 End With
            
            Next Lrow

    Columns("K:M").Delete
        End With
    
    
    'Move screen to the top
    Range("A1").Activate
    Set RNG = Nothing
     
End Sub
 
