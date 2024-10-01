VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} findRecordForm 
   Caption         =   "SMS - Find Record"
   ClientHeight    =   4404
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6180
   OleObjectBlob   =   "findRecordForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "findRecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'ensure all variables are defined


' Find values in search term ---- completed
Private Sub search_btn_Click()
    Dim searchTable As Range
    Dim cell As Range
    Dim matchFound As Boolean
    
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    Set searchTable = Selection
    
    searchResults_cmb.Clear ' clear combo list
    
    Range("A2").Select
    
    matchFound = False
    
    search_tbx.Text = Trim(search_tbx.Text)
    
    If Len(search_tbx.Text) < 3 Then
        MsgBox search_lbl.Caption & " must be minimum of 3 characters", vbExclamation
        search_tbx.SetFocus
        Exit Sub
    End If
    
    For Each cell In searchTable
        If InStr(LCase(cell.value), LCase(search_tbx.value)) >= 1 Then
            searchResults_cmb.AddItem cell.value ' add matched values to searchResults
            matchFound = True
        End If
    Next cell
    
    If Not matchFound Then
        MsgBox search_tbx.value & " not found in SalesID or Property Adress."
    Else
        flashColor
    End If
    
End Sub


' flash color of search results cmb if results are found
Sub flashColor()
    Dim originalColor As Long
    Dim i As Integer

    originalColor = searchResults_cmb.BackColor
    
    searchResults_cmb.BackColor = RGB(255, 255, 0)
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    searchResults_cmb.BackColor = RGB(100, 255, 0)
End Sub

Private Sub search_tbx_Change()

End Sub

' Scroll to selected value ---- completed
Private Sub searchResults_cmb_Change()
    Dim foundRange As Range
    Dim searchValue As String
    
    searchValue = searchResults_cmb.value
    If searchValue <> "" Then
        Set foundRange = ActiveSheet.UsedRange.Find(What:=searchValue, _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole, _
                                                    MatchCase:=False, _
                                                    SearchFormat:=False)
        If Not foundRange Is Nothing Then
            foundRange.Select
        End If
    End If
End Sub

' View record button
Private Sub view_btn_Click()
    Dim cell As Range
    Dim lastRow As Long
    Dim isFound As Boolean
    
    lastRow = Cells(rows.Count, "A").End(xlUp).Row
    isFound = False
    
    For Each cell In Range("A2:B" & lastRow)
        If findRecordForm.searchResults_cmb.value = cell.value Then
              Call showRecord(findRecordForm.searchResults_cmb.value)
              isFound = True
              Exit For
        End If
    Next cell
        
    If Not isFound Then
        MsgBox ("Record not found in database")
    End If
    
End Sub

' Clear records button
Private Sub clear_btn_Click()
    search_tbx.Text = ""
    searchResults_cmb.Clear
    searchResults_cmb.BackColor = RGB(255, 255, 255)
End Sub

' Exit form
Private Sub cancel_Click()
    Unload Me
End Sub

' Function to load record from search results combobox
Private Function showRecord(id As String)
    Dim entry As Range
    Dim entryList() As New Collection
    Dim i As Integer
    
    modifyRecordForm.ID_tbx = id
    modifyRecordForm.saveRecord_btn.Enabled = False

    modifyRecordForm.address_tbx.Text = Range("B" & ActiveCell.Row)
    modifyRecordForm.address_tbx.Locked = True
    modifyRecordForm.city_tbx.Text = Range("C" & ActiveCell.Row)
    modifyRecordForm.region_tbx.Text = Range("D" & ActiveCell.Row)
    modifyRecordForm.squareMeter_tbx.Text = Range("E" & ActiveCell.Row)
    modifyRecordForm.acreage_tbx.Text = Range("F" & ActiveCell.Row)
    modifyRecordForm.askingPrice_tbx.Text = Range("G" & ActiveCell.Row)
    modifyRecordForm.salesPrice_tbx.Text = Range("H" & ActiveCell.Row)
    modifyRecordForm.date_tbx.Text = Range("I" & ActiveCell.Row)
    modifyRecordForm.Show
   
    'showRecord
    
End Function



