Attribute VB_Name = "Isotopic_Warning"
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 10 May 2018
    'Looks for isotoptic analysis done on RCRA metals
    'Allows combination of some to a single elemental value
    
Sub RCRA_Isotopes()
Dim i As Integer
Dim RNG As Range
Dim ALNumber As Long
Dim ISOTOPE As Long
Dim MINTOMAX As Long
Dim CalcMode As Long
Dim ViewMode As Long
Dim Sample_No As Integer
Dim Answer As Integer
Dim RNG_1 As Range
Dim RNG_2 As Range
Dim Hello As Range
Dim Found As String

i = 2
    
    'Turn off Sceen Updating Make (single screen update at the end of the caluclation)
        With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = True
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
            
            'This section looks for and allows the user to input the total isotpoic mass of Ag as an elemental mass
            If Cells(i, "D") = "107Ag" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*Ag"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    AgTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Ag. Delete all other Ag Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If AgTotal = "False" Then End
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=AgTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Ag", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
        
            'This section looks for and allows the user to input the total isotpoic mass of Ag as an elemental mass
            If Cells(i, "D") = "109Ag" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*Ag"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    AgTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Ag. Delete all other Ag Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If AgTotal = "False" Then End
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=AgTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Ag", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of As as an elemental mass
            If Cells(i, "D") = "175As" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*As"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    AsTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for As. Delete all other As Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If AsTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=AsTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="As", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Be as an elemental mass
            If Cells(i, "D") = "9Be" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="*Be"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    BeTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Be. Delete all other Be Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If BeTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=BeTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Be", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Cd as an elemental mass
            If Cells(i, "D") = "111Cd" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*Cd"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    CdTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Cd. Delete all other Cd Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                    
                       If CdTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=CdTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Cd", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                                                            
            End If
            
            If Cells(i, "D") = "112Cd" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*Cd"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    CdTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Cd. Delete all other Cd Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If CdTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=CdTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Cd", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            If Cells(i, "D") = "113Cd" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*Cd"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    CdTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Cd. Delete all other Cd Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If CdTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=CdTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Cd", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            If Cells(i, "D") = "114Cd" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*Cd"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    CdTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Cd. Delete all other Cd Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If CdTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=CdTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Cd", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            If Cells(i, "D") = "116Cd" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="1*Cd"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    CdTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Cd. Delete all other Cd Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If CdTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=CdTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Cd", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Cr as an elemental mass
            If Cells(i, "D") = "52Cr" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="5*Cr"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    CrTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Cr. Delete all other Cr Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If CrTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=CrTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Cr", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Hg as an elemental mass
            If Cells(i, "D") = "202Hg" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="2*Hg"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    HgTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Hg. Delete all other Hg Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If HgTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=HgTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Hg", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Ni as an elemental mass
            If Cells(i, "D") = "58Ni" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="5*Ni"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    NiTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Ni. Delete all other Ni Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If NiTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=NiTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Ni", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Pb as an elemental mass
            If Cells(i, "D") = "208Pb" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="2*Pb"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    PbTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Pb. Delete all other Pb Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If PbTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=PbTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Pb", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Se as an elemental mass
            If Cells(i, "D") = "82Se" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="8*Se"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    SeTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Se. Delete all other Se Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If SeTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=SeTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Se", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
            
            'This section looks for and allows the user to input the total isotpoic mass of Tl as an elemental mass
            If Cells(i, "D") = "205Tl" Then
                Set RNG = Nothing
                    .AutoFilterMode = False
                    .Range("A1:H1").AutoFilter Field:=1, Criteria1:=ALNumber
                    .Range("A1:H1").AutoFilter Field:=2, Criteria1:=Sample_No
                    .Range("A1:H1").AutoFilter Field:=4, Criteria1:="2*Tl"
                Set RNG = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                    TlTotal = Application.InputBox("Isotope detected for AL#" & ALNumber & " Sample# " & Sample_No & "! Examine the units for the displayed isotopes,. If all the same sum to give a total value for Tl. Delete all other Tl Isotopes for AL#" & ALNumber & " Sample#" & Sample_No & "!", Title:="RCRA Isotpe Detected", Type:=1)
                       If TlTotal = "False" Then End
                       
                       Set RNG_1 = Range("E2:E" & Cells(Rows.Count, "E").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_1.Replace What:=Cells(i, "E").Value2, Replacement:=TlTotal, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                        Set RNG_2 = Range("D2:D" & Cells(Rows.Count, "D").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
                            RNG_2.Replace What:=Cells(i, "D"), Replacement:="Tl", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                            
            End If
                  
        i = i + 1
          
       Loop
    
    End With
        
End Sub

