Attribute VB_Name = "Module1"
Sub DeleteBlanks()

Application.CutCopyMode = True
Application.ScreenUpdating = False

Dim lRow As Integer
Dim intCol As Long
Dim rngCell As Range, fn
     
Set fn = Application.WorksheetFunction

For intCol = 8 To 8
    For lRow = 80 To 1 Step -1
        Set rngCell = Cells(lRow, intCol)
        With rngCell
            .Value = fn.Substitute(rngCell.Value, Chr(160), Chr(32))
            .Value = Trim(rngCell.Value)
           End With
        If Len(rngCell) = 0 Then
            rngCell.Delete shift:=xlUp
        End If
        Set rngCell = Nothing
    Next lRow
Next intCol
    
Application.CutCopyMode = False
Application.ScreenUpdating = True

End Sub

