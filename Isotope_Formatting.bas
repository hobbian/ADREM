Attribute VB_Name = "Isotope_Formatting"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine Written by Ian M. Hobbs 15 May 2018
'Subroutine Converst all the non-actinide isotopes to a single format

Sub Change_Species_Format_Iso()
Dim Lastrow As Long
Dim Firstrow As Long
Dim Lrow As Long
Dim Frow As Integer
Dim ViewMode As Long
Dim CalcMode As Long
 
    'Looks for data in worksheet labled Raw Data
    With Sheets("Raw Data")
        'Selects the sheet
        .Select
        'Changes the view to normal view
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView
        'Turn off Page Braks for Speed
        .DisplayPageBreaks = False
           
            Firstrow = 2
            Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
            
                'We loop from Lastrow to Firstrow
                For Lrow = Lastrow To Firstrow Step -1
                
                    With Cells(Lrow, "D")
                        
                       If .Value2 = "Mn-54" Then Cells(Lrow, "D").Value2 = "54Mn"
                        If .Value2 = "Co-60" Then Cells(Lrow, "D").Value2 = "60Co"
                        If .Value2 = "Sr-90" Then Cells(Lrow, "D").Value2 = "90Sr"
                        If .Value2 = "Y-90" Then Cells(Lrow, "D").Value2 = "90Y"
                        If .Value2 = "M/Z-90" Then Cells(Lrow, "D").Value2 = "90 M/Z"
                        If .Value2 = "Tc-99" Then Cells(Lrow, "D").Value2 = "99Tc"
                        If .Value2 = "Ru/Rh-106" Then Cells(Lrow, "D").Value2 = "106Ru/Rh"
                        If .Value2 = "Sb-125" Then Cells(Lrow, "D").Value2 = "125Sb"
                        If .Value2 = "Ba-134" Then Cells(Lrow, "D").Value2 = "134Ba"
                        If .Value2 = "Cs-134" Then Cells(Lrow, "D").Value2 = "134Cs"
                        If .Value2 = "Cs-137" Then Cells(Lrow, "D").Value2 = "137Cs"
                        If .Value2 = "Ba-137" Then Cells(Lrow, "D").Value2 = "137Ba"
                        If .Value2 = "M/Z-137" Then Cells(Lrow, "D").Value2 = "137M/Z"
                        If .Value2 = "Ba-138" Then Cells(Lrow, "D").Value2 = "138Ba"
                        If .Value2 = "Ce-144" Then Cells(Lrow, "D").Value2 = "144Ce"
                        If .Value2 = "Ce/Pr-144" Then Cells(Lrow, "D").Value2 = "144Ce"
                        If .Value2 = "Eu-154" Then Cells(Lrow, "D").Value2 = "154Eu"
                        If .Value2 = "Eu-155" Then Cells(Lrow, "D").Value2 = "155Eu"
                        If .Value2 = "U-233" Then Cells(Lrow, "D").Value2 = "233U"
                        If .Value2 = "U-234" Then Cells(Lrow, "D").Value2 = "234U"
                        If .Value2 = "U-235" Then Cells(Lrow, "D").Value2 = "235U"
                        If .Value2 = "U-236" Then Cells(Lrow, "D").Value2 = "236U"
                        If .Value2 = "U-238" Then Cells(Lrow, "D").Value2 = "238U"
                        If .Value2 = "Np-237" Then Cells(Lrow, "D").Value2 = "237Np"
                        If .Value2 = "M/Z-238" Then Cells(Lrow, "D").Value2 = "238M/Z"
                        If .Value2 = "Pu-238" Then Cells(Lrow, "D").Value2 = "238Pu"
                        If .Value2 = "Pu-239" Then Cells(Lrow, "D").Value2 = "239Pu"
                        If .Value2 = "Pu-240" Then Cells(Lrow, "D").Value2 = "240Pu"
                        If .Value2 = "Pu-241" Then Cells(Lrow, "D").Value2 = "241Pu"
                        If .Value2 = "Pu-242" Then Cells(Lrow, "D").Value2 = "242Pu"
                        If .Value2 = "M/Z-241" Then Cells(Lrow, "D").Value2 = "241M/Z"
                        If .Value2 = "Am-241" Then Cells(Lrow, "D").Value2 = "241Am"
                        If .Value2 = "Am-243" Then Cells(Lrow, "D").Value2 = "243Am"
                        If .Value2 = "Cm-244" Then Cells(Lrow, "D").Value2 = "244Cm"
                                        
                    End With
                
                Next
        
    End With
        
End Sub
