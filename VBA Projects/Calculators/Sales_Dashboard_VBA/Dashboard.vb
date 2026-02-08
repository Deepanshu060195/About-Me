
Sub CreateTrendChart()

    Dim ch As ChartObject
    
    Set ch = Sheets("Dashboard").ChartObjects.Add( _
        Left:=50, Top:=180, Width:=700, Height:=300)
    
    With ch.Chart
        .SetSourceData Sheets("Pivot_Backend").Range("A20").CurrentRegion
        .ChartType = xlColumnClustered
        
        .SeriesCollection(2).ChartType = xlLine
        .SeriesCollection(2).AxisGroup = 2
        
        .HasTitle = True
        .ChartTitle.Text = "Sales Quantity and Amount Trend"
    End With

End Sub


Sub CreateProfitDonut()

    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim ch As ChartObject
    
    Set pc = CreatePivotCache
    Set pt = pc.CreatePivotTable( _
        TableDestination:=Sheets("Pivot_Backend").Range("G1"), _
        TableName:="Profit_By_Product")
    
    With pt
        .PivotFields("Product").Orientation = xlRowField
        .AddDataField .PivotFields("Profit"), "Profit", xlSum
    End With
    
    Set ch = Sheets("Dashboard").ChartObjects.Add( _
        Left:=50, Top:=500, Width:=350, Height:=300)
    
    With ch.Chart
        .SetSourceData pt.TableRange1
        .ChartType = xlDoughnut
        .HasTitle = True
        .ChartTitle.Text = "Total Profit by Product"
    End With

End Sub


Sub CreateSlicers()

    Dim pt As PivotTable
    Dim sc As SlicerCache
    Dim wsDash As Worksheet
    
    Set wsDash = Sheets("Dashboard")
    Set pt = Sheets("Pivot_Backend").PivotTables("Monthly_Trend")

    'Make sure Pivot is refreshed
    pt.RefreshTable

    '======================
    ' PRODUCT SLICER
    '======================
    Set sc = ThisWorkbook.SlicerCaches.Add(pt, "Product")
    sc.Slicers.Add wsDash, , "Product_Slicer", "Product", 800, 80, 150, 200

    '======================
    ' YEAR SLICER
    '======================
    Set sc = ThisWorkbook.SlicerCaches.Add(pt, "Year")
    sc.Slicers.Add wsDash, , "Year_Slicer", "Year", 1000, 80, 150, 200

    '======================
    ' MONTH SLICER
    '======================
    Set sc = ThisWorkbook.SlicerCaches.Add(pt, "Month")
    sc.Slicers.Add wsDash, , "Month_Slicer", "Month", 1000, 300, 150, 250

    '======================
    ' COUNTRY SLICER
    '======================
    Set sc = ThisWorkbook.SlicerCaches.Add(pt, "Country")
    sc.Slicers.Add wsDash, , "Country_Slicer", "Country", 800, 300, 150, 250

    MsgBox "Slicers Created Successfully", vbInformation

End Sub
Sub ConnectSlicersToAllPivots()

    Dim sc As SlicerCache
    Dim pt As PivotTable
    Dim pf As PivotField

    Application.ScreenUpdating = False

    For Each sc In ThisWorkbook.SlicerCaches

        'Slicer must already be connected to at least one pivot
        If sc.PivotTables.Count > 0 Then

            For Each pt In Sheets("Pivot_Backend").PivotTables

                'Check PivotCache compatibility
                If pt.PivotCache.Index = sc.PivotTables(1).PivotCache.Index Then

                    'Check if slicer field exists in pivot
                    On Error Resume Next
                    Set pf = pt.PivotFields(sc.SourceName)
                    On Error GoTo 0

                    If Not pf Is Nothing Then
                        On Error Resume Next
                        sc.PivotTables.AddPivotTable pt
                        On Error GoTo 0
                    End If

                    Set pf = Nothing
                End If

            Next pt

        End If

    Next sc

    Application.ScreenUpdating = True
    MsgBox "Slicers connected successfully (no errors)", vbInformation

End Sub

Sub Refresh_Click()

    Dim pc As PivotCache

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    'Refresh all PivotCaches (fastest & safest)
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    'MsgBox "All PivotTables refreshed successfully", vbInformation

End Sub
Sub ImportSalesData()

    Dim fd As FileDialog
    Dim srcWB As Workbook
    Dim srcWS As Worksheet
    Dim tgtWS As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim filePath As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Set tgtWS = Sheets("Sales Data")

    '---------------------------
    ' Select file
    '---------------------------
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Sales Data File"
        .Filters.Clear
        .Filters.Add "Excel / CSV Files", "*.xlsx;*.xls;*.csv"
        .AllowMultiSelect = False

        If .Show <> -1 Then GoTo ExitHandler
        filePath = .SelectedItems(1)
    End With

    '---------------------------
    ' Open source file silently
    '---------------------------
    Set srcWB = Workbooks.Open(filePath, ReadOnly:=True)
    Set srcWS = srcWB.Sheets(1)

    '---------------------------
    ' Find data range
    '---------------------------
    lastRow = srcWS.Cells(srcWS.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWS.Cells(1, srcWS.Columns.Count).End(xlToLeft).Column

    '---------------------------
    ' Clear old data (keep headers)
    '---------------------------
    tgtWS.Rows("2:" & tgtWS.Rows.Count).ClearContents

    '---------------------------
    ' Copy new data
    '---------------------------
    srcWS.Range(srcWS.Cells(2, 1), srcWS.Cells(lastRow, lastCol)).Copy
    tgtWS.Cells(2, 1).PasteSpecial xlPasteValues

    srcWB.Close False

    '---------------------------
    ' Refresh dashboard
    '---------------------------
    Call FormatDateAndMonthColumns
    Call Refresh_Click

ExitHandler:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Sales data imported successfully", vbInformation

End Sub
Sub FormatDateAndMonthColumns()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = Sheets("Sales Data")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    'Find last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    '-----------------------------
    ' DATE COLUMN (dd/mm/yyyy)
    '-----------------------------
    For i = 2 To lastRow
        If ws.Cells(i, "A").Value <> "" Then
            ws.Cells(i, "A").Value = CDate(ws.Cells(i, "A").Value)
        End If
    Next i

    ws.Columns("A").NumberFormat = "dd/mm/yyyy"

    '-----------------------------
    ' MONTH COLUMN (TEXT ONLY)
    '-----------------------------
    For i = 2 To lastRow
        If ws.Cells(i, "C").Value <> "" Then
            ws.Cells(i, "C").Value = "'" & Format(ws.Cells(i, "C").Value, "mmmm")
        End If
    Next i

    ws.Columns("C").NumberFormat = "@"

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    'MsgBox "Date & Month columns formatted successfully", vbInformation

End Sub

