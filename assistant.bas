Attribute VB_Name = "assistant"

Function toggleColorMode(ByVal form As Object, ByVal toggle As Object)
    With form
        If toggle.value Then ' Check if the toggle is checked
            .BackColor = RGB(0, 0, 90)
            .ForeColor = RGB(255, 255, 255)
            toggle.ForeColor = RGB(0, 0, 0)
            toggle.BackColor = RGB(255, 255, 255)
            toggle.Caption = "Light Mode"
        Else
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
            toggle.ForeColor = RGB(255, 255, 255)
            toggle.BackColor = RGB(0, 0, 90)
            toggle.Caption = "Dark Mode"
        End If
    End With
End Function


Sub loadCombos()
    ' Add regions to the city combo box
    For Each region In Array("Midlands", "South England", "North England")
        addRecordForm.region_cmb.AddItem region
        Next region
    
    ' Add cities to the city combo box
    For Each city In Array("Lemington Spa", "Birmingham", "London", "Essex", "Liverpool", _
                           "Reading", "Manchester", "Bristol", "Newcastle", "Middlesborough")
        addRecordForm.city_cmb.AddItem city
        Next city
End Sub

Function getCities() As Variant
    getCities = Array("Lemington Spa", "Birmingham", "London", "Essex", "Liverpool", _
                           "Reading", "Manchester", "Bristol", "Newcastle", "Middlesborough")
End Function

Function regionCityDict() As Object
    Dim regionCity As Object
    Set regionCity = CreateObject("Scripting.Dictionary")
    
    ' define regions and thier cities
    regionCity.Add "Midlands", Array("Birmingham", "Bristol", "Lemington Spa", "Liverpool")
    regionCity.Add "North England", Array("Manchester", "Middlesborough", "Newcastle")
    regionCity.Add "South England", Array("Essex", "London", "Reading")
    regionCityDict = regionCity
End Function


Sub formatMouseIcon(button_object As Object)
    button_object.MousePointer = 99
    button_object.MouseIcon = LoadPicture(ThisWorkbook.Path & "\Hand Cursor.ico")
End Sub
