Attribute VB_Name = "Module5"
Public Sub AutoFillVersion1()

Dim myArray() As Variant
Dim DataRange As Range
Dim cell As Range
Dim x As Long
Dim searchTerm As String
Dim temp As Range
Dim countNonBlank As Integer
Dim myRange As Range

Workbooks("Floor Plan Creator.xlsm").Activate

Set DataRange = Range("H:H")

'Loop through each cell in Range and store value in Array
For Each cell In DataRange.Cells
    If cell <> "" Then
        ReDim Preserve myArray(x)
        myArray(x) = cell.Value
        x = x + 1
    Else
        Exit For
    End If
Next cell
  
Application.CutCopyMode = True
Application.ScreenUpdating = False

'Determines if the amount of columns in the employeeDatabase file is populated enough to start being used instead of the built in list of names.
'Set myRange = Columns("A:A")
'countNonBlank = Application.WorksheetFunction.CountA(myRange)

Workbooks("Floor Plan Creator.xlsm").Activate
'---------------------------PLATINUM SOUTH----------------------------
If Range("A3").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Brown, Scott"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B3").Value = searchTerm
    Else
    End If
End If
'---------------------------PLATINUM NORTH----------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A4").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Branidis, Nick"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B4").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Gayle, Lyndon"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B5").Value = searchTerm
    Else
    End If
End If
'------------------------------STAND 3--------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A7").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Howe, Rachel"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B7").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Ristoff, Meagan"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B8").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Salvi, Shreya"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B9").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 5-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A10").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Cheng, James"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("10").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Chau, Cynric"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B11").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Ebbin, Theresa"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B12").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 6-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A14").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Waines, Dimitri"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B14").Value = searchTerm
    Else
    End If
''-------------------------------SLOT 2--------------------------------
    searchTerm = "Spencer, Jahneese"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B15").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 7-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A17").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Newell, Gary"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B17").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Chang, Ming"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B18").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Bhachhoo, Hardeep Singh"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B19").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 8-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A22").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Telesford, Dominic"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B22").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 9-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A24").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Harthan, Peta-Gaye"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B24").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Chapple, Jaqueline"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B25").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Joseph, Michelle"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B26").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 10-------------------------------
'--------------------------------SLOT 1--------------------------------
If Range("A28").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Bent, Christopher"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B28").Value = searchTerm
    Else
    End If
'--------------------------------SLOT 2--------------------------------
    searchTerm = "Kim, Hoon"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B29").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 12-------------------------------
'--------------------------------SLOT 1--------------------------------
If Range("A31").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "De Crescenzo, Helena"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B31").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Malinay, Bernice Bea"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B32").Value = searchTerm
    Else
    End If
End If
'-------------------------------HOGTOWN 100-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A34").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Castillo, John"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B34").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Santos, John Revolver"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B35").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Esquivel, John"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B36").Value = searchTerm
    Else
    End If
End If
'-------------------------------P101-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A38").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Antonoglou, Michael"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B38").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Faddoul, Iman"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B39").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Bravo, Brian"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B40").Value = searchTerm
    Else
    End If
End If
'-------------------------------P103-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("A43").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Bryce, Matthew"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B43").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Teka, Kerry"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("B44").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 13-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C3").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Elliott, Christine"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D3").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Bostanci, Zeynel"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D4").Value = searchTerm
    Else
    End If
End If
'-------------------------------STAND 15-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C6").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Hernandez, Ericka"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D6").Value = searchTerm
    Else
    End If
End If
'------------------------------STAND 16-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C9").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Witter, Tiffany"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D9").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Ejidra, Sophie"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D10").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Iqbal, Javaid"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D11").Value = searchTerm
    Else
    End If
End If
'------------------------------STAND 17-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C13").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Alfaro, Erizalde"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D13").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Moore, Kyle"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D14").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Atangan, Justin"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D15").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 4--------------------------------
    searchTerm = "Hamilton, Brandon"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D16").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 5--------------------------------
    searchTerm = "Brown, Jermaine"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D17").Value = searchTerm
    Else
    End If
End If
'------------------------------STAND 18-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C18").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Soriano Sosa, Hilda Maria"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D18").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Peter, Brenton"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D19").Value = searchTerm
    Else
    End If
End If
'--------------------------HOGTOWN 300-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C21").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Castillo, Mara Jessa"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D21").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Tobias, Brissca"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D22").Value = searchTerm
    Else
    End If
End If
'------------------------------P322-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C23").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Rodney, Sophia"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D23").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "St. Clair, Alexander"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D24").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Glasgow, Jc"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D25").Value = searchTerm
    Else
    End If
End If
'--------------------------------MEZZ B-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C27").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Cabral, Jeff"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D27").Value = searchTerm
    Else
    End If
End If
'--------------------------------MEZZ D-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C29").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Jones, Jason"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D29").Value = searchTerm
    Else
    End If
End If
'------------------------------PANTRY A-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C31").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Caranglan, Jeaner"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D31").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Dionora, Erwin"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D32").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 3--------------------------------
    searchTerm = "Cox, Tom"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D33").Value = searchTerm
    Else
    End If
End If
'------------------------------PANTRY B-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C34").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Colicchio, Remington"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D34").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Nassief, Walter"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D35").Value = searchTerm
    Else
    End If
End If
'------------------------------PANTRY C-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C37").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Khenrab, Khenrab"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D37").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Dhakpa, Ngawang"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D38").Value = searchTerm
    Else
    End If
End If
'------------------------------PANTRY D-------------------------------
'-------------------------------SLOT 1--------------------------------
If Range("C40").Interior.Color = RGB(220, 230, 241) Then
    searchTerm = "Argel, Julius"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D40").Value = searchTerm
    Else
    End If
'-------------------------------SLOT 2--------------------------------
    searchTerm = "Nyemina, Marcel"
    If UBound(Filter(myArray, searchTerm)) >= 0 And searchTerm <> "" Then
        With Intersect(Columns("H"), ActiveSheet.UsedRange)
            .Replace searchTerm, " ", xlPart
        End With
        Range("D41").Value = searchTerm
    Else
    End If
End If

Call DeleteBlanks

End Sub
