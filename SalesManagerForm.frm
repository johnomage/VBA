VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SalesManagerForm 
   Caption         =   "Sales Management System"
   ClientHeight    =   2760
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SalesManagerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SalesManagerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addRecordbtn_Click()
'    addRecordForm.addRecord_btn.MousePointer = 99
'    addRecordForm.addRecord_btn.MouseIcon = LoadPicture(ThisWorkbook.Path & "\Hand Cursor.ico")
    Call assistant.formatMouseIcon(addRecordForm.addRecord_btn)
    loadCombos
    addRecordForm.Show
End Sub


Private Sub date_lbl_Click()
    SalesManagerLauncher.ShowDateTime
End Sub

Private Sub findRecord_btn_Click()
    findRecordForm.Show
End Sub
'
'Private Sub modifyRecordbtn_Click()
'    modifyRecordForm.Show
'End Sub

Private Sub quit_btn_Click()
    SalesManagerLauncher.StopTimer
    Unload SalesManagerForm
End Sub

Sub loadCombos()
    ' Add regions to the city combo box
    For Each region In Array("Midlands", "North England", "South England")
        addRecordForm.region_cmb.AddItem region
        Next region
    
    ' Add cities to the city combo box
    For Each city In Array("Birmingham", "Bristol", "Essex", "Lemington Spa", "Liverpool", _
                           "London", "Manchester", "Middlesborough", "Newcastle", "Reading")
        addRecordForm.city_cmb.AddItem city
        Next city
        
'    Select Case addRecordForm.region_cmb.value
'        Case "Midlands"
'            For Each city In Array("Birmingham", "Bristol", "Lemington Spa", "Liverpool")
'                addRecordForm.city_cmb.AddItem city
'                Next city
'        Case "South England"
'            For Each city In Array("Essex", "London", "Reading")
'                addRecordForm.city_cmb.AddItem city
'                Next city
'        Case "North England"
'            For Each city In Array("Manchester", "Middlesborough", "Newcastle")
'                addRecordForm.city_cmb.AddItem city
'                Next city
'        Case Else
'            addRecordForm.city_cmb.Clear
'    End Select
    
End Sub

