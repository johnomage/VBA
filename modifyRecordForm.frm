VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} modifyRecordForm 
   Caption         =   "SMS - View Record"
   ClientHeight    =   8916.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7668
   OleObjectBlob   =   "modifyRecordForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "modifyRecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' toggle dark or light mode
Private Sub darkMode_tog_Click()
    Call assistant.toggleColorMode(modifyRecordForm, darkMode_tog)
End Sub

Private Sub deleteRecord_btn_Click()
    Dim yesNo As VbMsgBoxResult
    Dim selectedID As String
    Dim deletedRecordsSheet As Worksheet
    Dim lastRow As Long
    
    selectedID = ID_tbx.Text
    
    If Len(ID_tbx.Text) < 8 Then
        MsgBox (ID_lbl.Caption & " is empty or incorrect!")
        Exit Sub
    End If
    
    yesNo = MsgBox("Are you sure you want to delete property record with ID: " & selectedID & "?", _
    vbYesNo + vbCritical + vbQuestion, _
    "Confirm Delete: " & ID_tbx.Text)
    
    If yesNo = vbYes Then
        If ActiveCell.value = selectedID Or ActiveCell.value = address_tbx.Text Then
            ActiveCell.EntireRow.Delete
            MsgBox "Record with ID " & selectedID & " has been deleted."
        End If
    Else
        MsgBox "Deletion canceled."
    End If
    
    Set deletedRecordsSheet = ThisWorkbook.Worksheets("DeletedRecords")
    
    Range("A2").Select
    lastRow = deletedRecordsSheet.Cells(deletedRecordsSheet.rows.Count, 1).End(xlUp).Row + 1 ' = Range("A2", Selection.End(xlDown))
    
    ' log deleted records in hidden DeletedRecords sheet.
    With modifyRecordForm
        deletedRecordsSheet.Cells(lastRow, 1).value = .ID_tbx.Text                       ' Sales ID column
        deletedRecordsSheet.Cells(lastRow, 2).value = .address_tbx.value                 ' Property Address column
        deletedRecordsSheet.Cells(lastRow, 3).value = .city_tbx.value                    ' City column
        deletedRecordsSheet.Cells(lastRow, 4).value = .region_tbx.value                  ' Region column
        deletedRecordsSheet.Cells(lastRow, 5).value = CDec(.squareMeter_tbx.value)       ' Square Meters column case as decimal
        deletedRecordsSheet.Cells(lastRow, 6).value = CDec(.acreage_tbx.value)           ' Acreage column cast as decimal
        deletedRecordsSheet.Cells(lastRow, 7).value = CDec(.askingPrice_tbx.value)       ' Asking Price column cast as Decimal
        deletedRecordsSheet.Cells(lastRow, 8).value = CDec(.salesPrice_tbx.value)        ' Sales Price column cast as decimal
        deletedRecordsSheet.Cells(lastRow, 9).value = CDate(.date_tbx.Text)              ' Date column cast as date
    End With
    
    Unload modifyRecordForm
    
        
End Sub

' modify record button click
Private Sub modifyRecord_btn_Click()
    modifyRecordForm.Caption = "SMS - Modify Record"
    saveRecord_btn.Enabled = True
    
    ' set text boxes to edit mode
    modifyRecordForm.address_tbx.Locked = False
    modifyRecordForm.city_tbx.Locked = False
    modifyRecordForm.region_tbx.Locked = False
    modifyRecordForm.squareMeter_tbx.Locked = False
    modifyRecordForm.acreage_tbx.Locked = False
    modifyRecordForm.askingPrice_tbx.Locked = False
    modifyRecordForm.salesPrice_tbx.Locked = False
    modifyRecordForm.date_tbx.Locked = False
    
End Sub

' save changes made to form entries
Private Sub saveRecord_btn_Click()
    Dim yesNo As VbMsgBoxResult
    
    If Len(ID_tbx.Text) < 8 Then
        MsgBox (ID_lbl.Caption & " is empty or incorrect!")
    Else
        yesNo = MsgBox("Do you want to save record?", vbYesNo + vbQuestion, "Confirm Save")
        Select Case yesNo
        Case vbYes
            Range("B" & ActiveCell.Row) = modifyRecordForm.address_tbx.Text
            Range("C" & ActiveCell.Row) = modifyRecordForm.city_tbx.Text
            Range("D" & ActiveCell.Row) = modifyRecordForm.region_tbx.Text
            Range("E" & ActiveCell.Row) = modifyRecordForm.squareMeter_tbx.Text
            Range("F" & ActiveCell.Row) = modifyRecordForm.acreage_tbx.Text
            Range("G" & ActiveCell.Row) = modifyRecordForm.askingPrice_tbx.Text
            Range("H" & ActiveCell.Row) = modifyRecordForm.salesPrice_tbx.Text
            Range("I" & ActiveCell.Row) = modifyRecordForm.date_tbx.Text
            
            Debug.Print ("Record saved")
            
        Case vbNo
            Debug.Print ("Save cancelled")
        End Select
    End If
    
    saveRecord_btn.Enabled = False
    
End Sub

' quit form
Private Sub exit_btn_Click()
    Unload Me
End Sub

