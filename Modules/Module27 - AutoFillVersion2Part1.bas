Attribute VB_Name = "Module27"
'Issue #1: Determine how to allocate the first slot in each location, then move
'on to the second slot. - Solved.
'Issue #2: Determine what condition will be used to end the For Loop after determining all slots have been filled. - Solved - Utilization of counter to determine how many times loop will run.
'Issue #3: Determine how to encompass all locations under one (or more?) ranges. Part 2 is a separate module, and theremore it might be trickier to connect the two. - Solved - Will use its own separate variables.
'-------------------------------BUILD ARRAY OF NAMES-------------------------------
Sub AutoFillVersion2Part1()

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
Static passes As Integer
Dim counter As Integer

'Static xMfreqA As String, xSMfreqA As String
'Static xMfreqB As String, xSMfreqB As String, xTMfreqB As String
'Static xMfreqC As String, xSMfreqC As String, xTMfreqC As String
'Static xMfreqD As String, xSMfreqD As String, xTMfreqD As String, xFMfreqD As String
'Static xMfreqE As String, xSMfreqE As String, xTMfreqE As String
'Static xMfreqF As String, xSMfreqF As String, xTMfreqF As String, xFMfreqF As String, xFiMfreqF As String
'Static xMfreqG As String, xSMfreqG As String
'Static xMfreqH As String, xSMfreqH As String, xTMfreqH As String, xFMfreqH As String
'Static xMfreqI As String, xSMfreqI As String
'Static xMfreqJ As String, xSMfreqJ As String
'Static xMfreqK As String, xSMfreqK As String, xTMfreqK As String
'Static xMfreqL As String, xSMfreqL As String, xTMfreqL As String, xFMfreqL As String
'Static xMfreqM As String, xSMfreqM As String, xTMfreqM As String, xFMfreqM As String, xFiMfreqM As String
'Static xMfreqN As String, xSMfreqN As String

Application.CutCopyMode = True
Application.ScreenUpdating = False

Workbooks.Open FileName:= _
"C:\Users\KyleHollands\OneDrive\Work Projects\Floor Plan Creator\employeeDatabase.xlsx"
'Workbooks.Open FileName:= _
'"C:\Users\KyleHollands\OneDrive\Floor Plan Creator\employeeDatabase.xlsx"

Set flrPlnRange = Workbooks("Floor Plan Creator.xlsm").Worksheets("Floor Plan Creator")
Set empDataRange = Workbooks("employeeDatabase.xlsx").Worksheets("Sheet1")

Set DataRange = flrPlnRange.Range("H:H")
'Loop through each cell in Range and store value in Array
For Each cell In DataRange.Cells
    If cell <> "" Then: ReDim Preserve myArray(x): myArray(x) = cell.Value: x = x + 1: Else: Exit For
Next cell
  
'Determines if the amount of columns in the employeeDatabase file is populated enough to start being used instead of the built in list of names.
countNonBlank = Application.WorksheetFunction.CountA(empDataRange.Range("A:A"))
If countNonBlank < 20 And countNonBlank <> Empty Then
    Call Autofill_Version1
'---------------------------------------------------------------------
Else
    Set dic = CreateObject("scripting.dictionary")
    On Error Resume Next
'---------------------------------------------------------------------
passes = InputBox("Enter the amount of passes.")
For Each n In flrPlnRange.Range("B:B")
    counter = counter + 1
'---------------------------PLATINUM SOUTH----------------------------
If flrPlnRange.Range("A3").Interior.Color = RGB(220, 230, 241) Then
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B3") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("A:A")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqA And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqA = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'prefEmp = "Brown, Scott"
'If UBound(Filter(myArray, prefEmp)) >= 0 Then: xOutValue = prefEmp
'--------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B3").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------PLATINUM NORTH----------------------------
If flrPlnRange.Range("A4").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B4") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("B:B")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqB And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqB = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B4").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B4") <> "" And flrPlnRange.Range("B5") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("B:B")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqB And xCount >= xSMfreqB And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqB = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B5").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B4") <> "" And flrPlnRange.Range("B5") <> "" And flrPlnRange.Range("B6") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("B:B")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqB And xCount >= xTMfreqB And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqB = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B6").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------STAND 3--------------------------------
If flrPlnRange.Range("A7").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B7") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("C:C")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqC And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqC = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B7").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B7") <> "" And flrPlnRange.Range("B8") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("C:C")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqC And xCount >= xSMfreqC And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqC = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B8").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B7") <> "" And flrPlnRange.Range("B8") <> "" And flrPlnRange.Range("B9") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("C:C")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqC And xCount >= xTMfreqC And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqC = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B9").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------STAND 5--------------------------------
If flrPlnRange.Range("A10").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B10") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("D:D")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqD And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqD = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B10").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B10") <> "" And flrPlnRange.Range("B11") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("D:D")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqD And xCount >= xSMfreqD And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqD = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B11").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B10") <> "" And flrPlnRange.Range("B11") <> "" And flrPlnRange.Range("B12") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("D:D")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqD And xCount >= xTMfreqD And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqD = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B12").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B10") <> "" And flrPlnRange.Range("B11") <> "" And flrPlnRange.Range("B12") <> "" And flrPlnRange.Range("B13") = "" Then
'-------------------------------SLOT 4--------------------------------
For Each Rng In empDataRange.Range("D:D")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xTMfreqD And xCount >= xFMfreqD And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFMfreqD = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B13").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------STAND 6--------------------------------
If flrPlnRange.Range("A14").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B14") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("E:E")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqE And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqE = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B14").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B14") <> "" And flrPlnRange.Range("B15") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("E:E")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqE And xCount >= xSMfreqE And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqE = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B15").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
''---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("B14") <> "" And flrPlnRange.Range("B15") <> "" And flrPlnRange.Range("B16") = "" Then
''-------------------------------SLOT 3--------------------------------
'For Each Rng In empDataRange.Range("E:E")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xSMfreqE And xCount >= xTMfreqE And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xTMfreqE = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("B16").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'------------------------------STAND 7--------------------------------
If flrPlnRange.Range("A17").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B17") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("F:F")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqF And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqF = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B17").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B17") <> "" And flrPlnRange.Range("B18") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("F:F")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqF And xCount >= xSMfreqF And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqF = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B18").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B17") <> "" And flrPlnRange.Range("B18") <> "" And flrPlnRange.Range("B19") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("F:F")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqF And xCount >= xTMfreqF And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqF = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B19").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B17") <> "" And flrPlnRange.Range("B18") <> "" And flrPlnRange.Range("B19") <> "" And flrPlnRange.Range("B20") = "" Then
'-------------------------------SLOT 4--------------------------------
For Each Rng In empDataRange.Range("F:F")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xTMfreqF And xCount >= xFMfreqF And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFMfreqF = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B20").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B17") <> "" And flrPlnRange.Range("B18") <> "" And flrPlnRange.Range("B19") <> "" And flrPlnRange.Range("B20") <> "" And flrPlnRange.Range("B21") = "" Then
'-------------------------------SLOT 5--------------------------------
For Each Rng In empDataRange.Range("F:F")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xFMfreqF And xCount >= xFiMfreqF And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFiMfreqF = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B21").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------STAND 8-----------------------------------
If flrPlnRange.Range("A22").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B22") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("G:G")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqG And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqG = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B22").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
''---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("B22") <> "" And flrPlnRange.Range("B23") = "" Then
''-------------------------------SLOT 2--------------------------------
'For Each Rng In empDataRange.Range("G:G")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xMfreqG And xCount >= xSMfreqG And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xSMfreqG = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("B23").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'------------------------------STAND 9--------------------------------
If flrPlnRange.Range("A24").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B24") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("H:H")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqH And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqH = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B24").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B24") <> "" And flrPlnRange.Range("B25") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("H:H")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqH And xCount >= xSMfreqH And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqH = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B25").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B24") <> "" And flrPlnRange.Range("B25") <> "" And flrPlnRange.Range("B26") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("H:H")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqH And xCount >= xTMfreqH And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqH = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B26").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B24") <> "" And flrPlnRange.Range("B25") <> "" And flrPlnRange.Range("B26") <> "" And flrPlnRange.Range("B27") = "" Then
'-------------------------------SLOT 4--------------------------------
For Each Rng In empDataRange.Range("H:H")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xTMfreqH And xCount >= xFMfreqH And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFMfreqH = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B27").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------STAND 10--------------------------------
If flrPlnRange.Range("A28").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B28") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("I:I")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqI And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqI = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B28").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B28") <> "" And flrPlnRange.Range("B29") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("I:I")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqI And xCount >= xSMfreqI And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqI = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B29").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------SWEET TOOTH--------------------------------
If flrPlnRange.Range("A30").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B30") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("J:J")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqJ And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqJ = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B30").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
''---------------------------------------------------------------------
'ElseIf flrPlnRange.Range("B30") <> "" And flrPlnRange.Range("B31") = "" Then
''-------------------------------SLOT 2--------------------------------
'For Each Rng In empDataRange.Range("J:J")
'    xValue = Rng.Value
'    If xValue <> "" Then
'        dic(xValue) = dic(xValue) + 1
'        xCount = dic(xValue)
'        If xCount <= xMfreqJ And xCount >= xSMfreqJ And UBound(Filter(myArray, xValue)) >= 0 Then
'                With flrPlnRange.Range("H:H")
'                    Set DupRng = .Find(What:=xValue)
'                        If Not DupRng Is Nothing Then: xSMfreqJ = xCount: xOutValue = xValue
'                End With
'        End If
'    Else: Exit For
'    End If
'Next
''---------------------------------------------------------------------
'With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
'    Set DupRng = .Find(What:=xOutValue)
'        If DupRng Is Nothing Then: xOutValue = ""
'            Range("B31").Value = xOutValue
'            With Intersect(Columns("H"), ActiveSheet.UsedRange)
'                .Replace xOutValue, " ", xlPart: End With
'End With
'Exit For
End If
Next
End If
'------------------------------STAND 12--------------------------------
If flrPlnRange.Range("A32").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange.Range
    If flrPlnRange.Range("B32") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("K:K")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqK And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqK = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B32").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B32") <> "" And flrPlnRange.Range("B33") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("K:K")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqK And xCount >= xSMfreqK And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqK = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B33").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B32") <> "" And flrPlnRange.Range("B33") <> "" And flrPlnRange.Range("B34") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("K:K")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqK And xCount >= xTMfreqK And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqK = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B34").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'---------------------------HogTown 100-------------------------------
If flrPlnRange.Range("A35").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B35") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("L:L")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqL And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqL = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B35").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B35") <> "" And flrPlnRange.Range("B36") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("L:L")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqL And xCount >= xSMfreqL And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqL = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B36").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B35") <> "" And flrPlnRange.Range("B36") <> "" And flrPlnRange.Range("B37") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("L:L")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqL And xCount >= xTMfreqL And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqL = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B37").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B35") <> "" And flrPlnRange.Range("B36") <> "" And flrPlnRange.Range("B37") <> "" And flrPlnRange.Range("B38") = "" Then
'-------------------------------SLOT 4--------------------------------
For Each Rng In empDataRange.Range("L:L")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xTMfreqL And xCount >= xFMfreqL And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFMfreqL = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B38").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------P101--------------------------------
If flrPlnRange.Range("A39").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B39") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("M:M")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqM And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqM = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B39").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B39") <> "" And flrPlnRange.Range("B40") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("M:M")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqM And xCount >= xSMfreqM And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqM = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B40").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B39") <> "" And flrPlnRange.Range("B40") <> "" And flrPlnRange.Range("B41") = "" Then
'-------------------------------SLOT 3--------------------------------
For Each Rng In empDataRange.Range("M:M")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xSMfreqM And xCount >= xTMfreqM And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xTMfreqM = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B41").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B39") <> "" And flrPlnRange.Range("B40") <> "" And flrPlnRange.Range("B41") <> "" And flrPlnRange.Range("B42") = "" Then
'-------------------------------SLOT 4--------------------------------
For Each Rng In empDataRange.Range("M:M")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xTMfreqM And xCount >= xFMfreqM And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFMfreqM = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B42").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B39") <> "" And flrPlnRange.Range("B40") <> "" And flrPlnRange.Range("B41") <> "" And flrPlnRange.Range("B42") <> "" And flrPlnRange.Range("B43") = "" Then
'-------------------------------SLOT 5--------------------------------
For Each Rng In empDataRange.Range("M:M")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xFMfreqM And xCount >= xFiMfreqM And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xFiMfreqM = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B43").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next
End If
'------------------------------P103-----------------------------------
If flrPlnRange.Range("A44").Interior.Color = RGB(220, 230, 241) Then
dic.RemoveAll
'---------------------------------------------------------------------
For Each n2 In flrPlnRange
    If flrPlnRange.Range("B44") = "" Then
'-------------------------------SLOT 1--------------------------------
For Each Rng In empDataRange.Range("N:N")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount >= xMfreqN And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xMfreqN = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B44").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
'---------------------------------------------------------------------
ElseIf flrPlnRange.Range("B44") <> "" And flrPlnRange.Range("B45") = "" Then
'-------------------------------SLOT 2--------------------------------
For Each Rng In empDataRange.Range("N:N")
    xValue = Rng.Value
    If xValue <> "" Then
        dic(xValue) = dic(xValue) + 1
        xCount = dic(xValue)
        If xCount <= xMfreqN And xCount >= xSMfreqN And UBound(Filter(myArray, xValue)) >= 0 Then
                With flrPlnRange.Range("H:H")
                    Set DupRng = .Find(What:=xValue)
                        If Not DupRng Is Nothing Then: xSMfreqN = xCount: xOutValue = xValue
                End With
        End If
    Else: Exit For
    End If
Next
'---------------------------------------------------------------------
With Sheets("Floor Plan Creator").Range("H:H"): Workbooks("Floor Plan Creator.xlsm").Activate
    Set DupRng = .Find(What:=xOutValue)
        If DupRng Is Nothing Then: xOutValue = ""
            Range("B45").Value = xOutValue
            With Intersect(Columns("H"), ActiveSheet.UsedRange)
                .Replace xOutValue, " ", xlPart: End With
End With
Exit For
End If
Next

End If

Call AutoFillVersion2Part2

If counter = passes Then: Exit For

Next

Call DeleteBlanks

End If

Workbooks("employeeDatabase.xlsx").Close Savechanges:=True

Application.CutCopyMode = False
Application.ScreenUpdating = True

End Sub
