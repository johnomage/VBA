Attribute VB_Name = "SalesManagerLauncher"
'Sub RectangleRoundedCorners1_Click()
'    SalesManagerForm.date_lbl = Now
'    SalesManagerForm.Show
'End Sub
Private nextUpdateTime As Date

' continuously update the date_lbl with the current time
Sub ShowDateTime()
    SalesManagerForm.date_lbl.Caption = Format(Now, "dddd, dd-mm-yyyy hh:nn:ss")
    
    ' Schedule the next update in 1 second
    nextUpdateTime = Now + TimeValue("00:00:01")
    Application.OnTime nextUpdateTime, "ShowDateTime"
End Sub

' Call this to start the timer and display the form
Sub RectangleRoundedCorners1_Click()
    Call assistant.formatMouseIcon(SalesManagerForm.addRecordbtn)
    Call assistant.formatMouseIcon(SalesManagerForm.findRecord_btn)
    Call assistant.formatMouseIcon(SalesManagerForm.quit_btn)
    Call ShowDateTime
    SalesManagerForm.Show
End Sub

' Optional: Use this to stop the timer if needed (e.g., when the form is closed)
Sub StopTimer()
    ' Cancel the scheduled update if it exists
    On Error Resume Next
    Application.OnTime nextUpdateTime, "ShowDateTime", , False
    On Error GoTo 0
End Sub

Sub ScrollDown()
Attribute ScrollDown.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("A1").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.offset(1, 0).Select
    ActiveCell.Formula2R1C1 = InputBox("Enter value")
    ActiveCell.Activate
End Sub
