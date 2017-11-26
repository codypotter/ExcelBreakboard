Attribute VB_Name = "Module7"
Public Const cashierRange As String = "A3:F14"
Public Const caRange As String = "A16:F20"
Public Const bohRange As String = "A22:F23"
Public Const supeRange As String = "A25:F28"
Public Const leadershipRange As String = "A30:F32"
Public Const dailyNotesRange As String = "K2:K25"
Public Const auditsRange As String = "N3:O10"
Public Const loginsRange As String = "R3:X10"

Sub sortByInTime()
Attribute sortByInTime.VB_ProcData.VB_Invoke_Func = " \n14"
' Cody Potter
' sortByInTime Macro
' This macro selects userdata in the breaksheet field and sorts it based on the "start" column.

    Range(cashierRange).Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B3:B14") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(cashierRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range(caRange).Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B16:B20") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(caRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range(bohRange).Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B22:B23") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(bohRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range(supeRange).Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B25:B28") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(supeRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range(leadershipRange).Select
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range("B30:B32") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(leadershipRange)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A3").Select
End Sub

Sub clearBreakboardData()
Attribute clearBreakboardData.VB_ProcData.VB_Invoke_Func = " \n14"
' Cody Potter
' clearBreakboardData Macro
' This macro clears all userdata and temporary formatting from the active sheet.

If MsgBox("Are you sure? This will clear all breaks, marks, and notes!", vbYesNo) = vbNo Then Exit Sub

    Range(cashierRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(caRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(bohRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(supeRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(leadershipRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(dailyNotesRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(auditsRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(loginsRange).Select
    Selection.ClearContents
    With Selection.Font
        .Strikethrough = False
        .FontStyle = "Regular"
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A3").Select
End Sub

Sub markCallOut()
' Cody Potter
' markCallOut Macro
' This macro marks the selected cell with bold, italic, and strikethrough. This indicates that a team member has called out for their shift.

    With Selection.Font
        .FontStyle = "Bold Italic"
        .Strikethrough = True
    End With
End Sub

Sub breakDone()
' Cody Potter
' This macro marks a cell green, to indicate a team member has returned from their break.

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
        .PatternTintAndShade = 0
    End With
End Sub
