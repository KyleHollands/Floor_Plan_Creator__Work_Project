Attribute VB_Name = "Module3"
Sub ReverseName()

Dim Rng As Range
Dim WorkRng As Range
Dim Sign As String

On Error Resume Next

Set WorkRng = Range("B3:B45, D3:D26, D28:D43, H1:H100")
Sign = ", "

For Each Rng In WorkRng
    xValue = Rng.Value
    NameList = VBA.Split(xValue, Sign)
    If UBound(NameList) = 1 Then
        Rng.Value = NameList(1) + Sign + NameList(0)
    End If
Next

End Sub
