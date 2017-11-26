Attribute VB_Name = "Module9"
Private Function printBreaksheet(ByVal Weekday As String)
' Cody Potter
' Function printBreaksheet
' This function accepts the active worksheet as an argument, and prints it's breaksheet accordingly.

    Range("A1:F32").Select
    Selection.Copy
    Sheets("Printer Friendly").Select
    Range("A1:F1").Select
    ActiveSheet.Paste
    Sheets(Weekday).Select
    Range("J1:K25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Printer Friendly").Select
    Range("H1:I1").Select
    ActiveSheet.Paste
    Range("A1:I32").Select
    Application.CutCopyMode = False
    With Selection.Font
        .Name = "Calibri"
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    With ActiveSheet.PageSetup
        .BlackAndWhite = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
    Application.Dialogs(xlDialogPrint).Show
    
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Function

Public Sub printBreaks()
' Cody Potter
' printBreaks Macro
' This macro chooses the correct function argument based on the active sheet.

    If ActiveWorkbook.Worksheets("Monday") Is ActiveSheet Then
        printBreaksheet ("Monday")
    ElseIf ActiveWorkbook.Worksheets("Tuesday") Is ActiveSheet Then
        printBreaksheet ("Tuesday")
    ElseIf ActiveWorkbook.Worksheets("Wednesday") Is ActiveSheet Then
        printBreaksheet ("Wednesday")
    ElseIf ActiveWorkbook.Worksheets("Thursday") Is ActiveSheet Then
        printBreaksheet ("Thursday")
    ElseIf ActiveWorkbook.Worksheets("Friday") Is ActiveSheet Then
        printBreaksheet ("Friday")
    ElseIf ActiveWorkbook.Worksheets("Saturday") Is ActiveSheet Then
        printBreaksheet ("Saturday")
    ElseIf ActiveWorkbook.Worksheets("Sunday") Is ActiveSheet Then
        printBreaksheet ("Sunday")
    End If
End Sub

Sub upcomingBreak()
Attribute upcomingBreak.VB_ProcData.VB_Invoke_Func = " \n14"
' Cody Potter
' upcomingBreak Macro
' This Macro highlights a break with a yellow background. This marks a break as 'upcoming'.

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599963377788629
        .PatternTintAndShade = 0
    End With
End Sub
Sub highlightBreak()
Attribute highlightBreak.VB_ProcData.VB_Invoke_Func = " \n14"
' Cody Potter
' highlightBreak Macro
' This marco highlights a break in an orange color, to indicate that the break is in progress.

    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 4944363
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

