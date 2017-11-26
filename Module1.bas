Attribute VB_Name = "Module1"

Public Sub assignBreaks()

    If MsgBox("Are you sure? This will assign default breaks!", vbYesNo) = vbNo Then Exit Sub
    
    Dim startTime As Date
    Dim endTime As Date
    Dim shiftLength As Long
    Dim shiftMidpoint As Long
    Dim oneBreakTime As Date
    Dim oneLunchTime As Date
    Dim twoBreakTime As Date
    
    ActiveSheet.Range(cashierRange).Select
        processBreaks
    
    ActiveSheet.Range(caRange).Select
        processBreaks
    
    ActiveSheet.Range(bohRange).Select
        processBreaks
    
    ActiveSheet.Range(supeRange).Select
        processBreaks
        
    ActiveSheet.Range(leadershipRange).Select
        processBreaks
        
End Sub

Public Sub processBreaks()
    ActiveCell.Offset(-1, 0).Select
    ActiveCell.Offset(1, 1).Select
    Do While (ActiveCell.Value <> Empty)
        startTime = Selection.Value
        ActiveCell.Offset(0, 4).Select
        endTime = Selection.Value
        
        shiftLength = DateDiff("n", startTime, endTime)
        shiftMidpoint = (shiftLength / 2)
        
        If shiftLength <= 300 Then
            oneBreakTime = DateAdd("n", shiftMidpoint, startTime)
            ActiveCell.Offset(0, -3).Select
            Selection.Value = oneBreakTime
            ActiveCell.Offset(0, -1).Select
        ElseIf shiftLength > 300 And shiftLength <= 390 Then
            oneBreakTime = DateAdd("n", 120, startTime)
            ActiveCell.Offset(0, -3).Select
            Selection.Value = oneBreakTime
            oneLunchTime = DateAdd("n", -120, endTime)
            ActiveCell.Offset(0, 1).Select
            Selection.Value = oneLunchTime
            ActiveCell.Offset(0, -2).Select
        ElseIf shiftLength > 390 Then
            oneBreakTime = DateAdd("n", 120, startTime)
            ActiveCell.Offset(0, -3).Select
            Selection.Value = oneBreakTime
            oneLunchTime = DateAdd("n", shiftMidpoint, startTime)
            ActiveCell.Offset(0, 1).Select
            Selection.Value = oneLunchTime
            twoBreakTime = DateAdd("n", -120, endTime)
            ActiveCell.Offset(o, 1).Select
            Selection.Value = twoBreakTime
            ActiveCell.Offset(0, -3).Select
        End If
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Text = "Start" Then
        Exit Sub
    End If
    Loop
End Sub
