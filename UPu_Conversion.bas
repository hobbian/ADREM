Attribute VB_Name = "UPu_Conversion"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine Written by Ian M. Hobbs 15 May 2018
    'Converts the Actinde elements to a single format

Sub Change_Spieces_Format_UPu()
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
                        
                        If .Value2 = "Total U" Then Cells(Lrow, "D").Value2 = "U Total"
                            If .Value2 = "U-233" Then Cells(Lrow, "D").Value2 = "233U"
                            If .Value2 = "U-234" Then Cells(Lrow, "D").Value2 = "234U"
                            If .Value2 = "U-235" Then Cells(Lrow, "D").Value2 = "235U"
                            If .Value2 = "U-236" Then Cells(Lrow, "D").Value2 = "236U"
                            If .Value2 = "U-238" Then Cells(Lrow, "D").Value2 = "238U"
                        If .Value2 = "Total Pu" Then Cells(Lrow, "D").Value2 = "Pu Total"
                            If .Value2 = "Pu-238" Then Cells(Lrow, "D").Value2 = "238Pu"
                            If .Value2 = "Pu-239" Then Cells(Lrow, "D").Value2 = "239Pu"
                            If .Value2 = "Pu-240" Then Cells(Lrow, "D").Value2 = "240Pu"
                            If .Value2 = "Pu-241" Then Cells(Lrow, "D").Value2 = "241Pu"
                            If .Value2 = "Pu-242" Then Cells(Lrow, "D").Value2 = "242Pu"
                        If .Value2 = "Am-241" Then Cells(Lrow, "D").Value2 = "241Am"
                        If .Value2 = "Am-243" Then Cells(Lrow, "D").Value2 = "243Am"
                                        
                    End With
                
                Next
        
        End With
        
End Sub
