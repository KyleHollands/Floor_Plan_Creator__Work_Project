Attribute VB_Name = "Module2"
Public Sub TrimNames()

Dim n As Object
Dim counter As Integer
Dim wrkRange As Range

Application.CutCopyMode = True
Application.ScreenUpdating = False

Set flrPlnRange = Workbooks("Floor Plan Creator.xlsm").Worksheets("Floor Plan Creator")

For Each n In flrPlnRange.Range("B3:B44")
    counter = counter + 1
        If n = "" Then
        ElseIf n <> "" Then n = Mid(n, InStr(n, ",") + 2)
        ElseIf counter = 41 Then
            Exit For
        End If
Next

For Each n In flrPlnRange.Range("D3:D25")
    counter = counter + 1
        If n = "" Then
        ElseIf n <> "" Then n = Mid(n, InStr(n, ",") + 2)
        ElseIf counter = 22 Then
            Exit For
        End If
Next

For Each n In flrPlnRange.Range("D27:D42")
    counter = counter + 1
        If n = "" Then
        ElseIf n <> "" Then n = Mid(n, InStr(n, ",") + 2)
        ElseIf counter = 15 Then
            Exit For
        End If
Next


'WorkRng = Range("B4")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B4").Value = WorkRng
'
'WorkRng = Range("B5")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B5").Value = WorkRng
'
'WorkRng = Range("B6")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B6").Value = WorkRng
'
'WorkRng = Range("B7")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B7").Value = WorkRng
'
'WorkRng = Range("B8")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B8").Value = WorkRng
'
'WorkRng = Range("B9")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B9").Value = WorkRng
'
'WorkRng = Range("B10")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B10").Value = WorkRng
'
'WorkRng = Range("B11")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B11").Value = WorkRng
'
'WorkRng = Range("B12")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B12").Value = WorkRng
'
'WorkRng = Range("B13")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B13").Value = WorkRng
'
'WorkRng = Range("B14")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B14").Value = WorkRng
'
'WorkRng = Range("B15")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B15").Value = WorkRng
'
'WorkRng = Range("B16")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B16").Value = WorkRng
'
'WorkRng = Range("B17")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B17").Value = WorkRng
'
'WorkRng = Range("B18")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B18").Value = WorkRng
'
'WorkRng = Range("B19")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B19").Value = WorkRng
'
'WorkRng = Range("B20")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B20").Value = WorkRng
'
'WorkRng = Range("B21")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B21").Value = WorkRng
'
'WorkRng = Range("B22")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B22").Value = WorkRng
'
'WorkRng = Range("B23")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B23").Value = WorkRng
'
'WorkRng = Range("B24")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B24").Value = WorkRng
'
'WorkRng = Range("B25")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B25").Value = WorkRng
'
'WorkRng = Range("B26")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B26").Value = WorkRng
'
'WorkRng = Range("B27")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B27").Value = WorkRng
'
'WorkRng = Range("B28")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B28").Value = WorkRng
'
'WorkRng = Range("B29")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B29").Value = WorkRng
'
'WorkRng = Range("B30")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B30").Value = WorkRng
'
'WorkRng = Range("B31")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B31").Value = WorkRng
'
'WorkRng = Range("B32")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B32").Value = WorkRng
'
'WorkRng = Range("B33")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B33").Value = WorkRng
'
'WorkRng = Range("B34")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B34").Value = WorkRng
'
'WorkRng = Range("B35")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B35").Value = WorkRng
'
'WorkRng = Range("B36")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B36").Value = WorkRng
'
'WorkRng = Range("B37")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B37").Value = WorkRng
'
'WorkRng = Range("B38")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B38").Value = WorkRng
'
'WorkRng = Range("B39")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("B39").Value = WorkRng
'
'WorkRng = Range("D3")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D3").Value = WorkRng
'
'WorkRng = Range("D4")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D4").Value = WorkRng
'
'WorkRng = Range("D5")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D5").Value = WorkRng
'
'WorkRng = Range("D6")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D6").Value = WorkRng
'
'WorkRng = Range("D7")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D7").Value = WorkRng
'
'WorkRng = Range("D8")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D8").Value = WorkRng
'
'WorkRng = Range("D9")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D9").Value = WorkRng
'
'WorkRng = Range("D10")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D10").Value = WorkRng
'
'WorkRng = Range("D11")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D11").Value = WorkRng
'
'WorkRng = Range("D12")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D12").Value = WorkRng
'
'WorkRng = Range("D13")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D13").Value = WorkRng
'
'WorkRng = Range("D14")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D14").Value = WorkRng
'
'WorkRng = Range("D15")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D15").Value = WorkRng
'
'WorkRng = Range("D16")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D16").Value = WorkRng
'
'WorkRng = Range("D17")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D17").Value = WorkRng
'
'WorkRng = Range("D18")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D18").Value = WorkRng
'
'WorkRng = Range("D19")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D19").Value = WorkRng
'
'WorkRng = Range("D20")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D20").Value = WorkRng
'
'WorkRng = Range("D21")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D21").Value = WorkRng
'
'WorkRng = Range("D23")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D23").Value = WorkRng
'
'WorkRng = Range("D24")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D24").Value = WorkRng
'
'WorkRng = Range("D25")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D25").Value = WorkRng
'
'WorkRng = Range("D26")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D26").Value = WorkRng
'
'WorkRng = Range("D27")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D27").Value = WorkRng
'
'WorkRng = Range("D28")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D28").Value = WorkRng
'
'WorkRng = Range("D29")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D29").Value = WorkRng
'
'WorkRng = Range("D30")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D30").Value = WorkRng
'
'WorkRng = Range("D31")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D31").Value = WorkRng
'
'WorkRng = Range("D32")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D32").Value = WorkRng
'
'WorkRng = Range("D33")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D33").Value = WorkRng
'
'WorkRng = Range("D34")
'WorkRng = Mid(WorkRng, InStr(WorkRng, ",") + 2)
'Range("D34").Value = WorkRng

Application.CutCopyMode = False
Application.ScreenUpdating = True

End Sub


