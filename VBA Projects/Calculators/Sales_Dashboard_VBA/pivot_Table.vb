Function CreatePivotCache() As PivotCache

    Dim srcData As Range
    Set srcData = Sheets("Sales Data").Range("A1").CurrentRegion
    
    Set CreatePivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcData)

End Function
Sub CreateKPIPivot()

    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Set pc = CreatePivotCache
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=Sheets("Pivot_Backend").Range("A1"), _
        TableName:="KPI_Pivot")
    
    With pt
        .AddDataField .PivotFields("Quantity"), "Total Quantity", xlSum
        .AddDataField .PivotFields("Sales"), "Total Sales", xlSum
        .AddDataField .PivotFields("Profit"), "Total Profit", xlSum
    End With

End Sub

Sub CreateMonthlyTrendPivot()
    Dim pc As PivotCache
    Dim pt As PivotTable
    
    Set pc = CreatePivotCache
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=Sheets("Pivot_Backend").Range("A20"), _
        TableName:="Monthly_Trend")
    
    With pt
        .PivotFields("Month").Orientation = xlRowField
        .AddDataField .PivotFields("Quantity"), "Quantity Sold", xlSum
        .AddDataField .PivotFields("Sales"), "Sales ($)", xlSum
    End With

End Sub



