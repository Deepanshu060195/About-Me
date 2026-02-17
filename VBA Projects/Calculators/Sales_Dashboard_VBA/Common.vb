Sub CreateDashboardSheets()

    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '-----------------------------
    ' Sales Data Sheet
    '-----------------------------
    If Not SheetExists("Sales Data") Then
        Set ws = Sheets.Add(after:=Sheets(Sheets.Count))
        ws.Name = "Sales Data"
    End If

    '-----------------------------
    ' Dashboard Sheet
    '-----------------------------
    If Not SheetExists("Dashboard") Then
        Set ws = Sheets.Add(before:=Sheets(1))
        ws.Name = "Dashboard"
    End If

    '-----------------------------
    ' Pivot Backend Sheet
    '-----------------------------
    If Not SheetExists("Pivot_Backend") Then
        Set ws = Sheets.Add(after:=Sheets(Sheets.Count))
        ws.Name = "Pivot_Backend"
        ws.Visible = xlSheetHidden
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Sheets checked / created successfully", vbInformation

End Sub
Function SheetExists(sheetName As String) As Boolean

    Dim ws As Worksheet
    SheetExists = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws

End Function

