Attribute VB_Name = "Activity_To_Mass_Conversion_Sub"


'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs Februrary 19, 2018
'Modified by Ian M. Hobbs May 8, 2018
    'Modification added all isotopes with specific activies
    
Sub Activity_To_Mass_Conversion()
Dim i As Long


'Sets the Row to begin the loop
i = 2

    'Selects Appropriate Sheet
    With Sheets("Raw Data")
        
        'Ensure only Rows with Data are Analyzed
        Do While Not IsEmpty(Sheets("Raw Data").Cells(i, "A"))
                                   
            'Uses specific activity data stored in Sample Totals Sheet to convert activity to grams
            If Cells(i, "D") = "233U" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(2, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)

            ElseIf Cells(i, "D") = "234U" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(3, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                
            ElseIf Cells(i, "D") = "235U" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(4, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
              
            ElseIf Cells(i, "D") = "236U" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(5, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                                  
            ElseIf Cells(i, "D") = "238U" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(6, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "238Pu" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(7, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "239Pu" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(8, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "240Pu" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(9, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "241Pu" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(10, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                                  
            ElseIf Cells(i, "D") = "242Pu" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(11, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "237Np" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(12, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "241Am" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(13, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                
            ElseIf Cells(i, "D") = "99Tc" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(14, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                         
            ElseIf Cells(i, "D") = "54Mn" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(15, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "60Co" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(16, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                       
            ElseIf Cells(i, "D") = "90Sr" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(17, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "90Y" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(18, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "106Ru/Rh" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(19, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "Sb" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(21, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                
            ElseIf Cells(i, "D") = "125Sb" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(21, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "134Cs" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(22, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                
            ElseIf Cells(i, "D") = "137Cs" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(23, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "144Ce" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(24, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "154Eu" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(25, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "155Eu" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(26, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                
            ElseIf Cells(i, "D") = "243Am" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(27, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
            
            ElseIf Cells(i, "D") = "244Cm" And Cells(i, "F") = "Ci" Then
                Cells(i, "F") = "g"
                Cells(i, "E").Value = Cells(i, "E").Value / Sheets("Lists").Cells(29, "I").Value2
                Cells(i, "E").Interior.Color = RGB(30, 186, 200)
                           
            End If
        
        i = i + 1
        
        Loop
        
    End With
                  
End Sub

