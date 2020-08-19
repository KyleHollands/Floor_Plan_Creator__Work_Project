Attribute VB_Name = "Module9"
Option Explicit

Sub CopyNames()

Dim Wkb As Worksheet

Set Wkb = Workbooks("Floor Plan Creator.xlsm").Worksheets("Floor Plan Creator")

Range("I:I").Copy Range("H:H")

End Sub
