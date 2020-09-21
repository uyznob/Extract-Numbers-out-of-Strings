Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit
Public icol As Integer
Public ocol As Integer
'n is the ID of input column (Ex: 1,2,3..)
'm is the ID of output coulmn (Ex: 4,5,6..)

'Purpose: Loop through input column
Sub Main()

Dim rng As Range
Dim i As Integer
Dim temp As String

MessBox.Show

Set rng = Range(Cells(1, icol), Cells(1, icol).End(xlDown))
For i = 1 To rng.Rows.Count
    temp = GetNumber(Cells(i, icol).Value)
    Cells(i, ocol).Value = CStr(temp)
Next

End Sub
