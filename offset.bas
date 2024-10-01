Attribute VB_Name = "offset"
Sub matrix()
    Dim rows As Integer
    Dim cols As String
    Dim i As Integer
    Dim j As Integer
    
    rows = CInt(InputBox("Enter number of rows"))
    cols = CInt(InputBox("Enter number of columns"))
    
    
    If ActiveCell.Row > 1 And ActiveCell.Column > 1 Then
        For i = 1 To rows
            For j = 1 To cols
            ActiveCell.offset(i, j).value = Round(WorksheetFunction.RandBetween(-50, 50) / 371.6, 4)
            Next j
        Next i
    End If
End Sub


Sub alastrow()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    MsgBox (lastRow)
End Sub
