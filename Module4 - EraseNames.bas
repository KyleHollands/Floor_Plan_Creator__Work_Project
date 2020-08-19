Attribute VB_Name = "Module4"
Public Sub EraseNames()

Dim Rng As Range
Dim WorkRng As Range
Dim Sign As String

On Error Resume Next

Set WorkRng = Range("B3:B45, D3:D26, D28:D43")

For Each Rng In WorkRng
    xValue = Rng.Value
    If UBound(NameList) = 1 Then
        Rng.Value = ""
    End If
Next

Call CopyNames

End Sub

