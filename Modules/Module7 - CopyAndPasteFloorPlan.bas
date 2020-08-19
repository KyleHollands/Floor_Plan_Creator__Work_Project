Attribute VB_Name = "Module7"
Public Sub CopyAndPasteFloorPlan()

Dim WBOpen

Range("A2:D45").Copy

WBOpen = IsWorkBookOpen("\\int.mlsel.com\1dept\FoodBeverage\Quick Service\03-QS New K Drive\Event Floor Plans - QS\Scotia Bank Arena\F20\Floor Plans - December 2019.xlsx")

If WBOpen = True Then

        Workbooks("Floor Plans - December 2019").Activate
        ActiveSheet.Range("A2").Select
        ActiveSheet.Paste
    
    Else
        Workbooks.Open FileName:= _
        "\\int.mlsel.com\1dept\FoodBeverage\Quick Service\03-QS New K Drive\Event Floor Plans - QS\Scotia Bank Arena\F20\Floor Plans - December 2019.xlsx"
        
        Workbooks("Floor Plans - December 2019").Activate
        ActiveSheet.Range("A2").Select
        ActiveSheet.Paste
    
    End If
    
End Sub

Function IsWorkBookOpen(FileName As String)

    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
    
End Function
