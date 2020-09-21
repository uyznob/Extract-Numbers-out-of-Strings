'Purpose: extract number in a cell
Private Function GetNumber(CellRef As String) As String

Dim StringLength As Integer
Dim i As Integer
Dim Result As String

StringLength = Len(CellRef)
For i = 1 To StringLength
    If (IsNumeric(Mid(CellRef, i, 1)) Or (Mid(CellRef, i, 1) = ".")) Then
        Result = Result & Mid(CellRef, i, 1)
    End If
Next i
GetNumber = Result

End Function
