Attribute VB_Name = "Module8"
Option Explicit

Sub CloseWorkbook()

Dim Wkb As Workbook

Set Wkb = Workbooks("Floor Plan Creator.xlsm")

Call UpdateDatabase
    
Wkb.Close Savechanges:=True

End Sub
