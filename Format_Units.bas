Attribute VB_Name = "Format_Units"
Option Explicit
'Copyright 2018, Battelle Energy Alliance, LLC  All Rights Reserved
'Subroutine written by Ian M. Hobbs 14 May 2018
    'Change the varied formats used in LIMS to a single format

Sub Unit_Formatting()
    
    'Make sure that Units are the same throughout the column
    'Changes formating on wt/wt concentrations to match
    Columns("F:F").Select
        Selection.Replace What:="ug/grams", Replacement:="µg/g", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="ug/g", Replacement:="µg/g", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="µg/grams", Replacement:="µg/g", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi/grams", Replacement:="µCi/g", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi/g", Replacement:="µCi/g", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="µCi/grams", Replacement:="µCi/g", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    'Changes formating on gamma spectroscopy concentrations to match
        Columns("F:F").Select
        Selection.Replace What:="µg/fractio", Replacement:="µg", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="ug/fractio", Replacement:="µg", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="µCi/sample", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="µCi/smpl", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi/sample", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi/samp", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False

    Columns("F:F").Select
        Selection.Replace What:="µCi/samp", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi/smpl", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi/spl", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="µCi/spl", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi/Smear", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="µCi/Smear", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="uCi", Replacement:="µCi", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="ug", Replacement:="µg", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="NA", Replacement:="n/a", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    Columns("F:F").Select
        Selection.Replace What:="N/A", Replacement:="n/a", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
                
End Sub
