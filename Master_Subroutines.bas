Attribute VB_Name = "Master_Subroutines"
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Macro written by J. Charboneau 22 March 18
'Subroutine modified by Ian M. Hobbs May 09 2018
    'Addition of all subroutines needed
    'Addition of the reinitiation commands was added by Ian M. Hobbs May 10, 2018
'Ctrl+Shift+M

Sub Master()
Dim Answer As Integer
Dim Answer_1 As Integer
Dim Answer_2 As Integer
Dim Answer_3 As Integer

    'Allows the user to choose into which subroutine the will eneter in the macro based on built in exit point in the macro
    Answer = MsgBox("Has macro previously been run and was aborted", vbYesNo + vbQuestion, "Reinitiation")
        If Answer = vbYes Then
            Answer_1 = MsgBox("Was the Macro aborted prior to deleting elemental and isotpics not tracked for waste purposes", vbYesNo + vbQuestion, "Reinitilize")
                
                If Answer_1 = vbYes Then GoTo ISO:
                If Answer_1 = vbNo Then
                    Answer_2 = MsgBox("Was the Macro aborted to enter Spl. Wt. data for unit conversion", vbYesNo + vbQuestion, "Reinitilize")
                End If
                
                        If Answer_2 = vbYes Then GoTo SplWt:
                        If Answer_2 = vbNo Then
                            Answer_3 = MsgBox("Was the Macro aborted prior to removing the non-conservative value", vbYesNo + vbQuestion, "Reinitilize")
                        End If
                        
                                If Answer_3 = vbYes Then GoTo RNCV:
                                If Answer_3 = vbNo Then
                                    MsgBox "Incorrect selection made reinitiate Macro and choose correct option", vbExclimation, "Unfortunate :("
                                End If
        End If
        
            'Perfomsmultiple subroutines sequentially accouding to call out number
                Call Backup
                Call Delete_Blanks
                Call Remove_Duplicates
                Call Unit_Formatting
                Call Format_Units_2
                Call Change_Spieces_Format_UPu
                Call Change_Species_Format_Iso
                Call RCRA_Isotopes
                Call DNED
ISO:            Call DEI
                Call Less_Than_Removal
SplWt:          Call WTPercent_Conversion
                Call ActivityConc_Conversion
                Call WeightConc_Conversion
                Call Activity_To_Mass_Conversion
                Call ISO_Conversion
RNCV:           Call Remove_Non_Conservative_Value
                Call Unit_Confirm
                    MsgBox "Macro has sucessfully been completed with no errors :)", vbOkay + vbExclamation, "Macro Complete"

End Sub

'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
' J. Charboneau 22 March 2018
' Backup Macro
' Create Backup of Raw Data

Sub Backup()
Dim i As Integer
Dim Answer As Integer

    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "Backup of Raw Data" Then
            Answer = MsgBox("Backup of Raw Data Exists. Delete the tab and re-run macro to create a new backup.;)", vbOkay + vbExclamation, "Backup Exists")
            End
        End If
    Next i
    

    Sheets("Raw Data").Select
    With Sheets("Raw Data")
        Cells.Select
        Selection.Copy
    End With
        
     Sheets.Add.Name = "Backup of Raw Data"
       
        Sheets("Backup of Raw Data").Select
        ActiveSheet.Paste

End Sub

'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine Written by Ian M. Hobbs 02 April 2018
'Ctrl+Shift+S

Sub Sample_Master()
Attribute Sample_Master.VB_ProcData.VB_Invoke_Func = "S\n14"
Dim CalcMode As Long
Dim Answer As Integer

Answer = MsgBox("Does the Consolidation Contain any dilutions", vbYesNo + vbQuestion, "Initiate Macro :0")
    If Answer = vbNo Then GoTo Sample:
    
        Call Transfer_Dwt_Dilution
        Call Dilution
        Call Transfer_DF
Sample: Call Transfer_Dwt
        Call Transfer_Data
        Call Sample_Totals
        Call Format
            MsgBox "All analytes have succssfully been calculated by the Macro. Verify all data input by user is correct prior to submission of data", vbOkay + vbExclamation, "Macro Complete"
            
  
     
End Sub
