Attribute VB_Name = "DNED_DEI"
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Ian M. Hobbs Februray 15, 2018
'Subroutine modified by Ian M. Hobbs 26 March 2018
'Subroutine modified to include RCRA metals by Ian M. Hobbs 03 April 2018
'Subroutine modified to Find and Replace varoius input units to a single format
    'by Ian M. Hobbs 07 May 2018
    
Sub DNED()
Dim Lrow As Long
Dim Firstrow As Long
Dim Lastrow As Long
Dim ViewMode As Long
Dim CalcMode As Long
Dim Amount As Double
Dim RNG As Range
Dim i As Long
Dim ISOTOPE As Variant
Dim MEASURE As Variant
Dim Measurement As Variant
Dim Deleted As Boolean
Dim Unit As Variant
Dim LResult As String
Dim Answer As Variant

counter = 1
i = 2

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

'Requires user to aknowledge they are ready to initiate macro
Answer = MsgBox("Do you wish to remove all data entries that have been canceled and/or have no results reported for analyses? Depending on amount of data calculation could take up to ten minutes do not exit program even if it says not responding.", vbYesNo + vbQuestion, "Initiate Macro :0")
    If Answer = vbNo Then End

                                
        'Set the First and Last row to loop through
        Firstrow = 2
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
            'We loop from Lastrow to Firstrow
            For Lrow = Lastrow To Firstrow Step -1
                
                'Chcek the values in the F column for gross alpha and deletes
                With .Cells(Lrow, "F")
                            
                        If .Value = "DPM/smear" Or .Value = "DPM/Smear" Then .EntireRow.Delete
                                           
                 End With
                 
            Next Lrow
            
        'Set the First and Last row to loop through
        Firstrow = 2
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
            'We loop from Lastrow to Firstrow
            For Lrow = Lastrow To Firstrow Step -1
                
                'Chcek the values in the E column for less thans and
                'removes the < and keeps the maximum value
                With .Cells(Lrow, "E")
                            
                        'Highlights values that are less than values
                         If .Value2 Like "<*" Then
                            Cells(Lrow, "I").Value2 = "Value was Converted from a <Value"
                            Cells(Lrow, "I").Interior.Color = RGB(240, 150, 250)
                         End If
                       
                    .Value2 = Replace(Cells(Lrow, "E").Value2, "<", "")
                   
                 End With
            
            Next Lrow

               
        'Set the First and Last row to loop through
        Firstrow = 2
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
        
            'We loop from Lastrow to Firstrow
            For Lrow = Lastrow To Firstrow Step -1
                
                'Removes lines without units
                With .Cells(Lrow, "F")
                            
                      If .Value2 = "n/a" Then .EntireRow.Delete
                       
                 End With
            
            Next Lrow

        End With
               
'Sets the Row to begin the loop
i = 2
    
    'Selects Appropriate Sheet
    With Sheets("Raw Data")
        
        'Ensure only Rows with Data are Analyzed
        Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
            Unit = Sheets("Raw Data").Cells(i, "F")
                                
            'Converts µCito total Ci
            If Unit = "µCi" Then
                Cells(i, "F") = "Ci"
                Cells(i, "E").Value = Cells(i, "E").Value * 0.000001
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
            
            'Converts µCi/mL to total Ci/g
            ElseIf Unit = "µCi/mL" Then
                Cells(i, "F") = "Ci/g"
                Cells(i, "E").Value = Cells(i, "E").Value * 0.000001 * 1.2
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
            
            
            'Converts DPM to total Ci
            ElseIf Unit = "DPM" Then
                Cells(i, "F") = "Ci"
                Cells(i, "F").Interior.Color = RGB(45, 176, 200)
                Cells(i, "E").Value = Cells(i, "E").Value / (2.22 * 10 ^ 12)
                
            'Converts µg to total g
            ElseIf Unit = "µg" Then
                Cells(i, "F") = "g"
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
                Cells(i, "E").Value = Cells(i, "E").Value * 0.000001
            

            'Converts mg to g
            ElseIf Unit = "mg" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value2 = Cells(i, "E").Value2 / 1000
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
                
            'Converts ng to g
            ElseIf Unit = "ng" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value2 = Cells(i, "E").Value2 * 0.000000001
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
                            
            'Converts the Wt% reported by ICP-OES to µg/g
            ElseIf Unit = "Wt%" And Cells(i, "C") = "ICP OES" Then
                Cells(i, "F") = "µg/g"
                Cells(i, "E").Value2 = Cells(i, "E").Value2 * 1000
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
             
            'Convets mg/g to µg/g
            ElseIf Unit = "mg/g" Then
                Cells(i, "F") = "µg/g"
                Cells(i, "E").Value2 = Cells(i, "E").Value2 / 1000
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
            
            'Convets µg/mL to µg/g
            ElseIf Unit = "µg/mL" Then
                Cells(i, "F") = "µg/g"
                Cells(i, "E").Value2 = Cells(i, "E").Value2 * 1.2
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
                
            'If the Cell is already in grams it is Highlighted
            ElseIf Unit = "g" Then
                Cells(i, "F").Interior.Color = RGB(45, 176, 240)
            
            End If
           
           'Resets the row to one down
            i = i + 1
       
        'Restarts the loop
        Loop

        
    End With
    
End Sub
     
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine deletes all isotopes and elements not of interest in waste disposal as of the FY18
'Subroutine Written by Ian M. Hobbs 09 May 2018

Sub DEI()
Dim Lrow As Long
Dim Firstrow As Long
Dim Lastrow As Long
Dim ViewMode As Long
Dim CalcMode As Long
Dim Amount As Double
Dim RNG As Range
Dim i As Long
Dim ISOTOPE As Variant
Dim MEASURE As Variant
Dim Measurement As Variant
Dim Deleted As Boolean
Dim Unit As Variant
Dim LResult As String
Dim Answer As Variant


counter = 1
i = 2

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
        
Answer = MsgBox("Do you want to delete all elements and isotopes not tracked for waste purposes", vbYesNo + vbExclamation, "Delete Unneeded Infromatin")
    If Answer = vbNo Then End
    
'Resets the i Value to 2
i = 2
     
    'Selects the isotopes relevant for tracking
    'Modified data will not be deleted due to replicate ElseIf statements
    With Sheets("Raw Data")
    
        Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
        
                ISOTOPE = Sheets("Raw Data").Cells(i, "D")
                MEASURE = Sheets("Raw Data").Cells(i, "C")
            Deleted = False
            
            If MEASURE = "Physical Measurements" Then GoTo nxt
                          
                'RCRA metals
                If ISOTOPE = "Ag" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "As" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Ba" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Be" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Cd" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Cr" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Hg" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Ni" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Pb" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Se" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Sb" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Tl" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                     
                'Isotopics
                ElseIf ISOTOPE = "54Mn" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "60Co" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "90Sr" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "90Y" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "90M/Z" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "106Ru/Rh" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "125Sb" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "134Cs" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "136Ba" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "137Cs" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "137Ba" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "137M/Z" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "138Ba" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "144Ce" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "154Eu" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "155Eu" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "241Am" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "243Am" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "237Np" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "99Tc" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "233U" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "234U" Then
                   Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "235U" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "236U" Then
                   Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "238U" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "238M/Z" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "238Pu" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "239Pu" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "240Pu" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "241Pu" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "241Am" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "241M/Z" Then
                     Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "242Pu" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "243Am" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "244Cm" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "Pu Total" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
                ElseIf ISOTOPE = "U Total" Then
                    Cells(i, "D").Interior.Color = RGB(159, 248, 110)
    
                Else
                    Rows(i).EntireRow.Delete
                    Deleted = True
                    
                End If
            
            If Not Deleted Then
            
nxt:               i = i + 1
            
            End If
            
        Loop
        
    End With
    
    End With
    
End Sub

