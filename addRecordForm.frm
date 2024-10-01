VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addRecordForm 
   Caption         =   "SMS - Add Record"
   ClientHeight    =   9456.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7380
   OleObjectBlob   =   "addRecordForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "addRecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub addRecord_btn_Click()
    Dim salesID As String 'AlpaheNumeric ID
    Dim region As String
    Dim city As String
    Dim dateSale As Date
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim salesPrice As Double
    
    ' data and entry validation
    With addRecordForm
        If .address_tbx.Text = "" Then
            MsgBox (.address_lbl.Caption & " cannot be empty!")
            Exit Sub
        ElseIf .region_cmb.value = "Select Region" Then
            MsgBox ("Select region in " & .region_lbl.Caption)
            Exit Sub
        ElseIf .city_cmb.value = "Select City" Then
            MsgBox ("Select city in " & .city_lbl.Caption)
            Exit Sub
        ElseIf .squareMeter_tbx.value = "" Or Not IsNumeric(.squareMeter_tbx.value) Then
            MsgBox (.squareMeter_lbl.Caption & " cannot be empty or must be numeric!")
            Exit Sub
        ElseIf .acreage_tbx.Text = "" Or Not IsNumeric(.acreage_tbx.Text) Then
            MsgBox (.acreage_lbl.Caption & " cannot be empty or must be numeric!")
            Exit Sub
        ElseIf .askingPrice_tbx.Text = "" Or Not IsNumeric(.askingPrice_tbx.Text) Then
            MsgBox (.askingPrice_lbl.Caption & " cannot be empty or must be numeric!")
            Exit Sub
        ElseIf .salesPrice_tbx.Text = "" Or Not IsNumeric(.salesPrice_tbx.Text) Then
            MsgBox (.salesPrice_lbl.Caption & " cannot be empty or must be numeric!")
            Exit Sub
        End If
        
        ' confirm sales values
        If Not IsValidPrice(.salesPrice_tbx.Text, .city_cmb.value) Or Not IsValid(.askingPrice_tbx.Text, .city_cmb.Visible) Then
            Dim response As VbMsgBoxResult
            response = MsgBox("The sales price is unusually low. Is there a discount?", vbYesNo + vbQuestion, "Confirm Price")
            If response = vbNo Then
                MsgBox ("Ok. Enter correct price")
                Exit Sub
            End If
        End If
        
    End With
    
    
        
    
    Set ws = ActiveSheet ' ThisWorkbook.Sheets("HomeSalesData")
    
    Range("A2").Select
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1 ' = Range("A2", Selection.End(xlDown))
    
    With addRecordForm
        ws.Cells(lastRow, 4).value = .region_cmb.value                  ' Region column
        ws.Cells(lastRow, 5).value = CDec(.squareMeter_tbx.value)       ' Square Meters column case as decimal
        ws.Cells(lastRow, 6).value = CDec(.acreage_tbx.value)           ' Acreage column cast as decimal
        ws.Cells(lastRow, 7).value = CDec(.askingPrice_tbx.value)       ' Asking Price column cast as Decimal
        ws.Cells(lastRow, 8).value = CDec(.salesPrice_tbx.value)        ' Sales Price column cast as decimal
        
        salesID = generateID(.city_cmb.value)
        city = .city_cmb.value
        
        If isUnique(salesID, city) Then
            ws.Cells(lastRow, 1).value = salesID                        ' Sales ID column
        End If
        
        ' check if city is in region
        If isCityInRegion(.city_cmb.value, .region_cmb.value) Then
            ws.Cells(lastRow, 2).value = .address_tbx.value             ' Property Address column
            ws.Cells(lastRow, 3).value = .city_cmb.value                ' City column
        Else
            Exit Sub
        End If
        
        ' check date format
        If IsDate(.date_tbx.Text) Then
            ws.Cells(lastRow, 9).value = CDate(.date_tbx.Text)          ' Date column cast as date
        Else
            MsgBox ("Please enter valid date in format DD/MM/YYYY")
            Exit Sub
        End If
        
        ' if all validation passed then
        MsgBox ("Record added with ID: " & generateID(.city_cmb.value))
        
    End With
    
    
    'reset the entire form
    ResetForm
    ' load city and region values
    'assistant.loadCombos
    
End Sub

' Toggle dark and light modes
Private Sub darkMode_tog_Click()
    Call toggleColorMode(addRecordForm, darkMode_tog)
End Sub

' Generate random sales ID
Private Function generateID(city)
    Dim salesID As String
    salesID = Left(city, 3) & WorksheetFunction.RandBetween(10000, 99999)
    generateID = UCase(salesID)
End Function

' Reset form values
Private Sub ResetForm()
    ' Clears all the fields in the form
    
    With addRecordForm
        .region_cmb.value = "Select Region"  ' Clear the region combo box
        .address_tbx.Text = ""               ' Clear the address text box
        .city_cmb.value = "Select City"      ' Clear the city combo box
        .squareMeter_tbx.value = ""          ' Clear the square meter text box
        .acreage_tbx.value = 0               ' Clear the acreage text box
        .askingPrice_tbx.value = 0           ' Clear the asking price text box
        .salesPrice_tbx.value = 0            ' Clear the sales price text box
        .date_tbx.Text = "DD/MM/YYYY"        ' Clear the date text box
    End With
        
End Sub

Private Function isUnique(salesID As String, city As String) As Boolean
        If WorksheetFunction.CountIf(Columns("A:A"), salesID) > 0 Then
            isUnique = False
            salesID = generateID(city)
        Else
            isUnique = True
        End If
End Function

Private Function isCityInRegion(city As String, region As String) As Boolean
    Dim cities As Variant
    Dim i As Integer
    ' declare regions dict with cities as value
    Dim regionCityDict As Object
    Set regionCityDict = CreateObject("Scripting.Dictionary")
    
    ' define regions and thier cities
    regionCityDict.Add "Midlands", Array("Birmingham", "Bristol", "Lemington Spa", "Liverpool")
    regionCityDict.Add "North England", Array("Manchester", "Middlesborough", "Newcastle")
    regionCityDict.Add "South England", Array("Essex", "London", "Reading")
    
    ' initialise function
    isCityInRegion = False
    
    ' check if the region exist
    If regionCityDict.exists(region) Then
        ' get cities in region
        cities = regionCityDict(region)
        
        For i = LBound(cities) To UBound(cities)
            If cities(i) = city Then
                isCityInRegion = True
                Exit For
            Else
                MsgBox (city & " is not in " & region & ". Please check again")
                Exit For
            End If
        Next i
        
    Else
        MsgBox ("Unknown region: " & region)
        Exit Function
    End If
            
'        Select Case addRecordForm.region_cmb.Value
'            Case "Midlands"
'                If WorksheetFunction.Match(city, midlands, 1) Then
'                    isCityInRegion = True
'
'
'            Case "South England"
'                For Each city In Array("Essex", "London", "Reading")
'                    addRecordForm.city_cmb.AddItem city
'                    Next city
'            Case "North England"
'                For Each city In Array("Manchester", "Middlesborough", "Newcastle")
'                    addRecordForm.city_cmb.AddItem city
'                    Next city
'            Case Else
'                addRecordForm.city_cmb.Clear
'    End Select
End Function

' confirm price of property
Private Function IsValidPrice(price As Double, city As String) As Boolean
    Dim cities As Variant
    Dim i As Integer
    
    cities = assistant.getCities
    'For i = LBound(cities) To UBound(cities)
    If city = "London" And price < 60000 Then
        IsValidPrice = False
    Else
        IsValidPrice = True
End Function

' Quit addRecord Application
Private Sub quit_btn_Click()
    Unload Me
End Sub


