Attribute VB_Name = "Module6"
Sub AutoFillVersion2Part2()

Dim myArray() As Variant
Dim DataRange As Range
Dim cell As Range
Dim x As Long
Dim searchTerm As String
Dim countNonBlank As Integer
Dim myRange As Range
Dim prefEmp As String
Dim Rng As Range
Dim DupRng As Range
Dim WorkRng As Range
Dim n As Variant
Dim n2 As Variant

Static xMfreqO As String, xSMfreqO As String, xTMfreqO As String
Static xMfreqP As String, xSMfreqP As String
Static xMfreqQ As String, xSMfreqQ As String
Static xMfreqR As String, xSMfreqR As String, xTMfreqR As String, xFMfreqR As String
Static xMfreqS As String, xSMfreqS As String, xTMfreqS As String, xFMfreqS As String, xFiMfreqS As String
Static xMfreqT As String, xSMfreqT As String, xTMfreqT As String
Static xMfreqU As String, xSMfreqU As String
Static xMfreqV As String, xSMfreqV As String, xTMfreqV As String
Static xMfreqW As String, xSMfreqW As String
Static xMfreqX As String, xSMfreqX As String
Static xMfreqY As String, xSMfreqY As String, xTMfreqY As String
Static xMfreqZ As String, xSMfreqZ As String, xTMfreqZ As String
Static xMfreqAA As String, xSMfreqAA As String, xTMfreqAA As String
Static xMfreqAB As String, xSMfreqAB As String, xTMfreqAB As String

Application.CutCopyMode = True
Application.ScreenUpdating = False

Set flrPlnRange = Workbooks("Floor Plan Creator.xlsm").Worksheets("Floor Plan Creator")
Set empDataRange = Workbooks("employeeDatabase.xlsx").Worksheets("Sheet1")

Set DataRange = flrPlnRange.Range("H:H")
'Loop through each cell in Range and store value in Array
For Each cell In DataRange.Cells
    If cell <> "" Then: ReDim Preserve myArray(x): myArray(x) = cell.Value: x = x + 1: Else: Exit For
Next cell

Set dic = CreateObject("scripting.dictionary")
On Error Resume Next
'---------------------------STAND 13----------------------------
If flrPlnRange.Range("C3").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D3") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("O:O")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqO And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqO = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D3").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D3") <> "" And flrPlnRange.Range("D4") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("O:O")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqO And xCount >= xSMfreqO And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqO = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D4").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D3") <> "" And flrPlnRange.Range("D4") <> "" And flrPlnRange.Range("D5") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("O:O")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqO And xCount >= xTMfreqO And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqO = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D5").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------STAND 15-----------------------------------
If flrPlnRange.Range("C6").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D6") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("P:P")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqP And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqP = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D6").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
''---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("D6") <> "" And flrPlnRange.Range("D7") = "" Then
''-------------------------------SLOT 2--------------------------------
'For Each Rng In empDataRange.Range("P:P")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xMfreqP And xCount >= xSMfreqP And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xSMfreqP = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D7").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'---------------------------PANTRY E-----------------------------------
If flrPlnRange.Range("C8").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D8") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("Q:Q")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqQ And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqQ = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D8").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
''---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("D8") <> "" And flrPlnRange.Range("D9") = "" Then
''-------------------------------SLOT 2--------------------------------
'For Each Rng In empDataRange.Range("Q:Q")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xMfreqQ And xCount >= xSMfreqQ And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xSMfreqQ = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D9").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'---------------------------STAND 16-------------------------------
If flrPlnRange.Range("C10").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D10") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("R:R")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqR And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqR = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D10").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D10") <> "" And flrPlnRange.Range("D11") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("R:R")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqR And xCount >= xSMfreqR And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqR = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D11").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D10") <> "" And flrPlnRange.Range("D11") <> "" And flrPlnRange.Range("D12") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("R:R")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqR And xCount >= xTMfreqR And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqR = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D12").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D10") <> "" And flrPlnRange.Range("D11") <> "" And flrPlnRange.Range("D12") <> "" And flrPlnRange.Range("D13") = "" Then
'-------------------------------SLOT 4--------------------------------
For Each Rng In empDataRange.Range("R:R")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xTMfreqR And xCount >= xFMfreqR And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFMfreqR = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D13").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------STAND 17--------------------------------
If flrPlnRange.Range("C14").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D14") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("S:S")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqS And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqS = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D14").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D14") <> "" And flrPlnRange.Range("D15") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("S:S")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqS And xCount >= xSMfreqS And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqS = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D15").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D14") <> "" And flrPlnRange.Range("D15") <> "" And flrPlnRange.Range("D16") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("S:S")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqS And xCount >= xTMfreqS And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqS = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D16").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D14") <> "" And flrPlnRange.Range("D15") <> "" And flrPlnRange.Range("D16") <> "" And flrPlnRange.Range("D17") = "" Then
'-------------------------------SLOT 4--------------------------------
For Each Rng In empDataRange.Range("S:S")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xTMfreqS And xCount >= xFMfreqS And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFMfreqS = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D17").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D14") <> "" And flrPlnRange.Range("D15") <> "" And flrPlnRange.Range("D16") <> "" And flrPlnRange.Range("D17") <> "" And flrPlnRange.Range("D18") = "" Then
'-------------------------------SLOT 5--------------------------------
For Each Rng In empDataRange.Range("S:S")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xFMfreqS And xCount >= xFiMfreqS And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFiMfreqS = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D18").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------STAND 18----------------------------
If flrPlnRange.Range("C19").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D19") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("T:T")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqT And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqT = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D19").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D19") <> "" And flrPlnRange.Range("D20") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("T:T")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqT And xCount >= xSMfreqT And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqT = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D20").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D19") <> "" And flrPlnRange.Range("D20") <> "" And flrPlnRange.Range("21") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("T:T")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqT And xCount >= xTMfreqT And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqT = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D21").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------HT 300-----------------------------------
If flrPlnRange.Range("C22").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D22") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("U:U")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqU And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqU = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D22").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D22") <> "" And flrPlnRange.Range("D23") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("U:U")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqU And xCount >= xSMfreqU And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqU = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D23").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------P322----------------------------
If flrPlnRange.Range("C24").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D24") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("V:V")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqV And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqV = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D24").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D24") <> "" And flrPlnRange.Range("D25") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("V:V")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqV And xCount >= xSMfreqV And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqV = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D25").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D24") <> "" And flrPlnRange.Range("D25") <> "" And flrPlnRange.Range("D26") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("V:V")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqV And xCount >= xTMfreqV And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqV = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D26").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------MEZZ B-----------------------------------
If flrPlnRange.Range("C28").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D28") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("W:W")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqW And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqW = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D28").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D28") <> "" And flrPlnRange.Range("D29") = "" Then
''-------------------------------SLOT 2--------------------------------
'For Each Rng In empDataRange.Range("W:W")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xMfreqW And xCount >= xSMfreqW And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xSMfreqW = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D29").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'---------------------------MEZZ D-----------------------------------
If flrPlnRange.Range("C30").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D30") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("X:X")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqX And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqX = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D30").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
''---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("D30") <> "" And flrPlnRange.Range("D31") = "" Then
''-------------------------------SLOT 2--------------------------------
'For Each Rng In empDataRange.Range("X:X")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xMfreqX And xCount >= xSMfreqW And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xSMfreqX = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D31").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'---------------------------PANTRY A----------------------------
If flrPlnRange.Range("C32").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D32") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("Y:Y")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqY And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqY = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D32").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D32") <> "" And flrPlnRange.Range("D33") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("Y:Y")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqY And xCount >= xSMfreqY And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqY = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D33").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
''---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("D32") <> "" And flrPlnRange.Range("D33") <> "" And flrPlnRange.Range("D34") = "" Then
''-------------------------------SLOT 3--------------------------------
'For Each Rng In empDataRange.Range("Y:Y")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xSMfreqY And xCount >= xTMfreqY And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xTMfreqY = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D34").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'---------------------------PANTRY B----------------------------
If flrPlnRange.Range("C35").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D35") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("Z:Z")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqZ And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqZ = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D35").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D35") <> "" And flrPlnRange.Range("D36") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("Z:Z")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqZ And xCount >= xSMfreqZ And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqZ = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D36").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("D35") <> "" And flrPlnRange.Range("D36") <> "" And flrPlnRange.Range("D37") = "" Then
''-------------------------------SLOT 3--------------------------------
'For Each Rng In empDataRange.Range("Z:Z")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xSMfreqZ And xCount >= xTMfreqZ And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xTMfreqZ = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D37").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'---------------------------PANTRY C----------------------------
If flrPlnRange.Range("C38").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D38") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("AA:AA")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqAA And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqAA = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D38").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D38") <> "" And flrPlnRange.Range("D39") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("AA:AA")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqAA And xCount >= xSMfreqAA And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqAA = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D39").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("D38") <> "" And flrPlnRange.Range("D39") <> "" And flrPlnRange.Range("D40") = "" Then
''-------------------------------SLOT 3--------------------------------
'For Each Rng In empDataRange.Range("AA:AA")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xSMfreqAA And xCount >= xTMfreqAA And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xTMfreqAA = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D40").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'---------------------------PANTRY D----------------------------
If flrPlnRange.Range("C41").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("D41") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("AB:AB")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqAB And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqAB = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D41").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("D41") <> "" And flrPlnRange.Range("D42") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("AB:AB")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqAB And xCount >= xSMfreqAB And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqAB = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("D42").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("D41") <> "" And flrPlnRange.Range("D42") <> "" And flrPlnRange.Range("D43") = "" Then
''-------------------------------SLOT 3--------------------------------
'For Each Rng In empDataRange.Range("AB:AB")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xSMfreqAB And xCount >= xTMfreqAB And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xTMfreqAB = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("D43").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If

End Sub

