Attribute VB_Name = "Module29"
Public Sub UpdateDatabase()

Application.CutCopyMode = True
Application.ScreenUpdating = False

'------------------------------UPDATE DATABASE------------------------
Workbooks.Open FileName:= _
"C:\Users\khollands\Desktop\Floor Plan Creator (Updated)\employeeDatabase.xlsx"

Set flrPlnRange = Workbooks("Floor Plan Creator.xlsm").Worksheets("Floor Plan Creator")
Set empDataRange = Workbooks("employeeDatabase.xlsx").Worksheets("Sheet1")

    '--------------PLATINUM SOUTH-------------
    If flrPlnRange.Range("B3") <> "" Then
        flrPlnRange.Range("B3").Copy
        Lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        Range("A" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("A:A").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------PLATINUM NORTH-------------
    If flrPlnRange.Range("B4") <> "" Then
        flrPlnRange.Range("B4").Copy
        Lastrow = Cells(Rows.Count, "B").End(xlUp).Row
        Range("B" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("B:B").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B5") <> "" Then
        flrPlnRange.Range("B5").Copy
        Lastrow = Cells(Rows.Count, "B").End(xlUp).Row
        Range("B" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("B:B").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B6") <> "" Then
        flrPlnRange.Range("B6").Copy
        Lastrow = Cells(Rows.Count, "B").End(xlUp).Row
        Range("B" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("B:B").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 3-------------
    If flrPlnRange.Range("B7") <> "" Then
        flrPlnRange.Range("B7").Copy
        Lastrow = Cells(Rows.Count, "C").End(xlUp).Row
        Range("C" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("C:C").Sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B8") <> "" Then
        flrPlnRange.Range("B8").Copy
        Lastrow = Cells(Rows.Count, "C").End(xlUp).Row
        Range("C" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("C:C").Sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B9") <> "" Then
        flrPlnRange.Range("B9").Copy
        Lastrow = Cells(Rows.Count, "C").End(xlUp).Row
        Range("C" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("C:C").Sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 5-------------
    If flrPlnRange.Range("B10") <> "" Then
        flrPlnRange.Range("B10").Copy
        Lastrow = Cells(Rows.Count, "D").End(xlUp).Row
        Range("D" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("D:D").Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B11") <> "" Then
        flrPlnRange.Range("B11").Copy
        Lastrow = Cells(Rows.Count, "D").End(xlUp).Row
        Range("D" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("D:D").Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B12") <> "" Then
        flrPlnRange.Range("B12").Copy
        Lastrow = Cells(Rows.Count, "D").End(xlUp).Row
        Range("D" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("D:D").Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B13") <> "" Then
        flrPlnRange.Range("B13").Copy
        Lastrow = Cells(Rows.Count, "D").End(xlUp).Row
        Range("D" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("D:D").Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 6-------------
    If flrPlnRange.Range("B14") <> "" Then
        flrPlnRange.Range("B14").Copy
        Lastrow = Cells(Rows.Count, "E").End(xlUp).Row
        Range("E" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("E:E").Sort Key1:=Range("E1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B15") <> "" Then
        flrPlnRange.Range("B15").Copy
        Lastrow = Cells(Rows.Count, "E").End(xlUp).Row
        Range("E" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("E:E").Sort Key1:=Range("E1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B16") <> "" Then
        flrPlnRange.Range("B16").Copy
        Lastrow = Cells(Rows.Count, "E").End(xlUp).Row
        Range("E" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("E:E").Sort Key1:=Range("E1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 7-------------
    If flrPlnRange.Range("B17") <> "" Then
        flrPlnRange.Range("B17").Copy
        Lastrow = Cells(Rows.Count, "F").End(xlUp).Row
        Range("F" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("F:F").Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B18") <> "" Then
        flrPlnRange.Range("B18").Copy
        Lastrow = Cells(Rows.Count, "F").End(xlUp).Row
        Range("F" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("F:F").Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B19") <> "" Then
        flrPlnRange.Range("B19").Copy
        Lastrow = Cells(Rows.Count, "F").End(xlUp).Row
        Range("F" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("F:F").Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B20") <> "" Then
        flrPlnRange.Range("B20").Copy
        Lastrow = Cells(Rows.Count, "F").End(xlUp).Row
        Range("F" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("F:F").Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B21") <> "" Then
        flrPlnRange.Range("B21").Copy
        Lastrow = Cells(Rows.Count, "F").End(xlUp).Row
        Range("F" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("F:F").Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 8-------------
    If flrPlnRange.Range("B22") <> "" Then
        flrPlnRange.Range("B22").Copy
        Lastrow = Cells(Rows.Count, "G").End(xlUp).Row
        Range("G" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("G:G").Sort Key1:=Range("G1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B23") <> "" Then
        flrPlnRange.Range("B23").Copy
        Lastrow = Cells(Rows.Count, "G").End(xlUp).Row
        Range("G" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("G:G").Sort Key1:=Range("G1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 9-------------
    If flrPlnRange.Range("B24") <> "" Then
        flrPlnRange.Range("B24").Copy
        Lastrow = Cells(Rows.Count, "H").End(xlUp).Row
        Range("H" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("H:H").Sort Key1:=Range("H1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B25") <> "" Then
        flrPlnRange.Range("B25").Copy
        Lastrow = Cells(Rows.Count, "H").End(xlUp).Row
        Range("H" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("H:H").Sort Key1:=Range("H1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B26") <> "" Then
        flrPlnRange.Range("B26").Copy
        Lastrow = Cells(Rows.Count, "H").End(xlUp).Row
        Range("H" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("H:H").Sort Key1:=Range("H1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B27") <> "" Then
        flrPlnRange.Range("B27").Copy
        Lastrow = Cells(Rows.Count, "H").End(xlUp).Row
        Range("H" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("H:H").Sort Key1:=Range("H1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 10-------------
    If flrPlnRange.Range("B28") <> "" Then
        flrPlnRange.Range("B28").Copy
        Lastrow = Cells(Rows.Count, "I").End(xlUp).Row
        Range("I" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("I:I").Sort Key1:=Range("I1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B29") <> "" Then
        flrPlnRange.Range("B29").Copy
        Lastrow = Cells(Rows.Count, "I").End(xlUp).Row
        Range("I" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("I:I").Sort Key1:=Range("I1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------SWEET TOOTH-------------
    If flrPlnRange.Range("B30") <> "" Then
        flrPlnRange.Range("B30").Copy
        Lastrow = Cells(Rows.Count, "J").End(xlUp).Row
        Range("J" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("J:J").Sort Key1:=Range("J1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 12-------------
    If flrPlnRange.Range("B32") <> "" Then
        flrPlnRange.Range("B32").Copy
        Lastrow = Cells(Rows.Count, "K").End(xlUp).Row
        Range("K" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("K:K").Sort Key1:=Range("K1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B33") <> "" Then
        flrPlnRange.Range("B33").Copy
        Lastrow = Cells(Rows.Count, "K").End(xlUp).Row
        Range("K" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("K:K").Sort Key1:=Range("K1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B34") <> "" Then
        flrPlnRange.Range("B34").Copy
        Lastrow = Cells(Rows.Count, "K").End(xlUp).Row
        Range("K" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("K:K").Sort Key1:=Range("K1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------HOGTOWN 100-------------
    If flrPlnRange.Range("B35") <> "" Then
        flrPlnRange.Range("B35").Copy
        Lastrow = Cells(Rows.Count, "L").End(xlUp).Row
        Range("L" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("L:L").Sort Key1:=Range("L1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B36") <> "" Then
        flrPlnRange.Range("B36").Copy
        Lastrow = Cells(Rows.Count, "L").End(xlUp).Row
        Range("L" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("L:L").Sort Key1:=Range("L1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B37") <> "" Then
        flrPlnRange.Range("B37").Copy
        Lastrow = Cells(Rows.Count, "L").End(xlUp).Row
        Range("L" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("L:L").Sort Key1:=Range("L1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B38") <> "" Then
        flrPlnRange.Range("B38").Copy
        Lastrow = Cells(Rows.Count, "L").End(xlUp).Row
        Range("L" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("L:L").Sort Key1:=Range("L1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------P101-------------
    If flrPlnRange.Range("B39") <> "" Then
        flrPlnRange.Range("B39").Copy
        Lastrow = Cells(Rows.Count, "M").End(xlUp).Row
        Range("M" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("M:M").Sort Key1:=Range("M1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B40") <> "" Then
        flrPlnRange.Range("B40").Copy
        Lastrow = Cells(Rows.Count, "M").End(xlUp).Row
        Range("M" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("M:M").Sort Key1:=Range("M1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B41") <> "" Then
        flrPlnRange.Range("B41").Copy
        Lastrow = Cells(Rows.Count, "M").End(xlUp).Row
        Range("M" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("M:M").Sort Key1:=Range("M1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B42") <> "" Then
        flrPlnRange.Range("B42").Copy
        Lastrow = Cells(Rows.Count, "M").End(xlUp).Row
        Range("M" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("M:M").Sort Key1:=Range("M1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B43") <> "" Then
        flrPlnRange.Range("B43").Copy
        Lastrow = Cells(Rows.Count, "M").End(xlUp).Row
        Range("M" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("M:M").Sort Key1:=Range("M1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------P103-------------
    If flrPlnRange.Range("B44") <> "" Then
        flrPlnRange.Range("B44").Copy
        Lastrow = Cells(Rows.Count, "N").End(xlUp).Row
        Range("N" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("N:N").Sort Key1:=Range("N1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("B45") <> "" Then
        flrPlnRange.Range("B45").Copy
        Lastrow = Cells(Rows.Count, "N").End(xlUp).Row
        Range("N" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("N:N").Sort Key1:=Range("N1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------STAND 13-----------
    If flrPlnRange.Range("D3") <> "" Then
        flrPlnRange.Range("D3").Copy
        Lastrow = Cells(Rows.Count, "O").End(xlUp).Row
        Range("O" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("O:O").Sort Key1:=Range("O1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D4") <> "" Then
        flrPlnRange.Range("D4").Copy
        Lastrow = Cells(Rows.Count, "O").End(xlUp).Row
        Range("O" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("O:O").Sort Key1:=Range("O1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D5") <> "" Then
        flrPlnRange.Range("D5").Copy
        Lastrow = Cells(Rows.Count, "O").End(xlUp).Row
        Range("O" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("O:O").Sort Key1:=Range("O1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------STAND 15-----------
    If flrPlnRange.Range("D6") <> "" Then
        flrPlnRange.Range("D6").Copy
        Lastrow = Cells(Rows.Count, "P").End(xlUp).Row
        Range("P" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("P:P").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D7") <> "" Then
        flrPlnRange.Range("D7").Copy
        Lastrow = Cells(Rows.Count, "P").End(xlUp).Row
        Range("P" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("P:P").Sort Key1:=Range("P1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
        '------------Pantry E-----------
    If flrPlnRange.Range("D8") <> "" Then
        flrPlnRange.Range("D8").Copy
        Lastrow = Cells(Rows.Count, "Q").End(xlUp).Row
        Range("Q" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Q:Q").Sort Key1:=Range("Q1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D9") <> "" Then
        flrPlnRange.Range("D9").Copy
        Lastrow = Cells(Rows.Count, "Q").End(xlUp).Row
        Range("Q" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Q:Q").Sort Key1:=Range("Q1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 16-------------
    If flrPlnRange.Range("D10") <> "" Then
        flrPlnRange.Range("D10").Copy
        Lastrow = Cells(Rows.Count, "R").End(xlUp).Row
        Range("R" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("R:R").Sort Key1:=Range("R1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D11") <> "" Then
        flrPlnRange.Range("D11").Copy
        Lastrow = Cells(Rows.Count, "R").End(xlUp).Row
        Range("R" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("R:R").Sort Key1:=Range("R1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D12") <> "" Then
        flrPlnRange.Range("D12").Copy
        Lastrow = Cells(Rows.Count, "R").End(xlUp).Row
        Range("R" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("R:R").Sort Key1:=Range("R1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D13") <> "" Then
        flrPlnRange.Range("D13").Copy
        Lastrow = Cells(Rows.Count, "R").End(xlUp).Row
        Range("R" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("R:R").Sort Key1:=Range("R1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------STAND 17-------------
    If flrPlnRange.Range("D14") <> "" Then
        flrPlnRange.Range("D14").Copy
        Lastrow = Cells(Rows.Count, "S").End(xlUp).Row
        Range("S" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("S:S").Sort Key1:=Range("S1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D15") <> "" Then
        flrPlnRange.Range("D15").Copy
        Lastrow = Cells(Rows.Count, "S").End(xlUp).Row
        Range("S" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("S:S").Sort Key1:=Range("S1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D16") <> "" Then
        flrPlnRange.Range("D16").Copy
        Lastrow = Cells(Rows.Count, "S").End(xlUp).Row
        Range("S" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("S:S").Sort Key1:=Range("S1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D17") <> "" Then
        flrPlnRange.Range("D17").Copy
        Lastrow = Cells(Rows.Count, "S").End(xlUp).Row
        Range("S" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("S:S").Sort Key1:=Range("S1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D18") <> "" Then
        flrPlnRange.Range("D18").Copy
        Lastrow = Cells(Rows.Count, "S").End(xlUp).Row
        Range("S" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("S:S").Sort Key1:=Range("S1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------STAND 18-----------
    If flrPlnRange.Range("D19") <> "" Then
        flrPlnRange.Range("D19").Copy
        Lastrow = Cells(Rows.Count, "T").End(xlUp).Row
        Range("T" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("T:T").Sort Key1:=Range("T1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D20") <> "" Then
        flrPlnRange.Range("D20").Copy
        Lastrow = Cells(Rows.Count, "T").End(xlUp).Row
        Range("T" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("T:T").Sort Key1:=Range("T1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D21") <> "" Then
        flrPlnRange.Range("D21").Copy
        Lastrow = Cells(Rows.Count, "T").End(xlUp).Row
        Range("T" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("T:T").Sort Key1:=Range("T1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------HOGHOUSE 300-----------
    If flrPlnRange.Range("D22") <> "" Then
        flrPlnRange.Range("D22").Copy
        Lastrow = Cells(Rows.Count, "U").End(xlUp).Row
        Range("U" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("U:U").Sort Key1:=Range("U1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D23") <> "" Then
        flrPlnRange.Range("D23").Copy
        Lastrow = Cells(Rows.Count, "U").End(xlUp).Row
        Range("U" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("U:U").Sort Key1:=Range("U1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------P322-------------
    If flrPlnRange.Range("D24") <> "" Then
        flrPlnRange.Range("D24").Copy
        Lastrow = Cells(Rows.Count, "V").End(xlUp).Row
        Range("V" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("V:V").Sort Key1:=Range("V1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D25") <> "" Then
        flrPlnRange.Range("D25").Copy
        Lastrow = Cells(Rows.Count, "V").End(xlUp).Row
        Range("V" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("V:V").Sort Key1:=Range("V1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D26") <> "" Then
        flrPlnRange.Range("D26").Copy
        Lastrow = Cells(Rows.Count, "V").End(xlUp).Row
        Range("V" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("V:V").Sort Key1:=Range("V1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------MEZZ B-------------
    If flrPlnRange.Range("D28") <> "" Then
        flrPlnRange.Range("D28").Copy
        Lastrow = Cells(Rows.Count, "W").End(xlUp).Row
        Range("W" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("W:W").Sort Key1:=Range("W1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D29") <> "" Then
        flrPlnRange.Range("D29").Copy
        Lastrow = Cells(Rows.Count, "W").End(xlUp).Row
        Range("W" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("W:W").Sort Key1:=Range("W1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '--------------MEZZ D-------------
    If flrPlnRange.Range("D30") <> "" Then
        flrPlnRange.Range("D30").Copy
        Lastrow = Cells(Rows.Count, "X").End(xlUp).Row
        Range("X" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("X:X").Sort Key1:=Range("X1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D31") <> "" Then
        flrPlnRange.Range("D31").Copy
        Lastrow = Cells(Rows.Count, "X").End(xlUp).Row
        Range("X" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("X:X").Sort Key1:=Range("X1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------PANTRY A-----------
    If flrPlnRange.Range("D32") <> "" Then
        flrPlnRange.Range("D32").Copy
        Lastrow = Cells(Rows.Count, "Y").End(xlUp).Row
        Range("Y" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Y:Y").Sort Key1:=Range("Y1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D33") <> "" Then
        flrPlnRange.Range("D33").Copy
        Lastrow = Cells(Rows.Count, "Y").End(xlUp).Row
        Range("Y" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Y:Y").Sort Key1:=Range("Y1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D34") <> "" Then
        flrPlnRange.Range("D34").Copy
        Lastrow = Cells(Rows.Count, "Y").End(xlUp).Row
        Range("Y" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Y:Y").Sort Key1:=Range("Y1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------PANTRY B-----------
    If flrPlnRange.Range("D35") <> "" Then
        flrPlnRange.Range("D35").Copy
        Lastrow = Cells(Rows.Count, "Z").End(xlUp).Row
        Range("Z" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Z:Z").Sort Key1:=Range("Z1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D36") <> "" Then
        flrPlnRange.Range("D36").Copy
        Lastrow = Cells(Rows.Count, "Z").End(xlUp).Row
        Range("Z" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Z:Z").Sort Key1:=Range("Z1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D37") <> "" Then
        flrPlnRange.Range("D37").Copy
        Lastrow = Cells(Rows.Count, "Z").End(xlUp).Row
        Range("Z" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("Z:Z").Sort Key1:=Range("Z1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------PANTRY C-----------
    If flrPlnRange.Range("D38") <> "" Then
        flrPlnRange.Range("D38").Copy
        Lastrow = Cells(Rows.Count, "AA").End(xlUp).Row
        Range("AA" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("AA:AA").Sort Key1:=Range("AA1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D39") <> "" Then
        flrPlnRange.Range("D39").Copy
        Lastrow = Cells(Rows.Count, "AA").End(xlUp).Row
        Range("AA" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("AA:AA").Sort Key1:=Range("AA1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D40") <> "" Then
        flrPlnRange.Range("D40").Copy
        Lastrow = Cells(Rows.Count, "AA").End(xlUp).Row
        Range("AA" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("AA:AA").Sort Key1:=Range("AA1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    '------------PANTRY D-----------
    If flrPlnRange.Range("D41") <> "" Then
        flrPlnRange.Range("D41").Copy
        Lastrow = Cells(Rows.Count, "AB").End(xlUp).Row
        Range("AB" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("AB:AB").Sort Key1:=Range("AB1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D42") <> "" Then
        flrPlnRange.Range("D42").Copy
        Lastrow = Cells(Rows.Count, "AB").End(xlUp).Row
        Range("AB" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("AB:AB").Sort Key1:=Range("AB1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    If flrPlnRange.Range("D43") <> "" Then
        flrPlnRange.Range("D43").Copy
        Lastrow = Cells(Rows.Count, "AB").End(xlUp).Row
        Range("AB" & (Lastrow + 1)).PasteSpecial xlPasteValuesAndNumberFormats
        Columns("AB:AB").Sort Key1:=Range("AB1"), Order1:=xlAscending, Header:=xlYes
        Else: End If
    
    Workbooks("employeeDatabase.xlsx").Activate
    Workbooks("employeeDatabase.xlsx").Close Savechanges:=True
        
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
End Sub
