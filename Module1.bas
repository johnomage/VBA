Attribute VB_Name = "Module1"
Sub findNegative2D()
    Dim cell As Range
    Dim negativeCount As Integer
    
    negativeCount = 0
    
    For Each cell In Selection
        If cell.value < 0 Then
            negativeCount = negativeCount + 1
        End If
    Next cell
    
    Range("N2").value = negativeCount
    
    MsgBox ("Negative Count: " & negativeCount)
    
    Call scanNegatives
End Sub

Private Sub scanNegatives()
    Dim cell As Range
    
    For Each cell In Selection
        If cell.value < 0 Then
            cell.ClearContents
        End If
    Next cell
End Sub
