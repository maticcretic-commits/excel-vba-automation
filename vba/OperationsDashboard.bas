Attribute VB_Name = "OperationsDashboard"
Option Explicit

'======================================================================
' Operations Dashboard
' ---------------------------------------------------------------------
' Rebuilds the Dashboard sheet from raw rows on the Data sheet.
' Shows KPI tiles, a Top-5 table, conditional formatting, and two
' auto-generated charts. Re-runnable: wipes Dashboard first.
'
' Expected workbook layout:
'   Data      : A=Date | B=Category | C=Metric | D=Value
'   Dashboard : (built by this module)
'======================================================================

Public Sub RefreshDashboard()
    Dim dataWs As Worksheet, dashWs As Worksheet
    Set dataWs = ThisWorkbook.Worksheets("Data")
    Set dashWs = ThisWorkbook.Worksheets("Dashboard")

    Dim lastRow As Long
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Data sheet is empty.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    dashWs.Cells.Clear
    Dim ch As ChartObject
    For Each ch In dashWs.ChartObjects
        ch.Delete
    Next ch
    Application.DisplayAlerts = True

    '----- Title -------------------------------------------------------
    With dashWs.Range("B2:H2")
        .Merge
        .Value = "Operations Dashboard"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlLeft
    End With
    dashWs.Range("B3").Value = "Last refreshed: " & Format(Now, "dd-mmm-yyyy hh:mm")
    dashWs.Range("B3").Font.Italic = True
    dashWs.Range("B3").Font.Color = RGB(120, 120, 120)

    '----- KPI tiles ---------------------------------------------------
    Dim totalValue As Double, avgValue As Double, maxValue As Double
    Dim countRows As Long
    countRows = lastRow - 1
    totalValue = Application.WorksheetFunction.Sum(dataWs.Range("D2:D" & lastRow))
    avgValue = Application.WorksheetFunction.Average(dataWs.Range("D2:D" & lastRow))
    maxValue = Application.WorksheetFunction.Max(dataWs.Range("D2:D" & lastRow))

    WriteKpi dashWs, "B5", "Total", totalValue, RGB(31, 78, 121)
    WriteKpi dashWs, "D5", "Average", avgValue, RGB(84, 130, 53)
    WriteKpi dashWs, "F5", "Peak", maxValue, RGB(192, 80, 77)
    WriteKpi dashWs, "H5", "Records", countRows, RGB(128, 100, 162)

    '----- Top-5 table -------------------------------------------------
    Dim tableRow As Long
    tableRow = 8
    dashWs.Cells(tableRow, 2).Value = "Top 5 by Value"
    dashWs.Cells(tableRow, 2).Font.Bold = True
    dashWs.Cells(tableRow, 2).Font.Size = 14
    dashWs.Cells(tableRow + 1, 2).Resize(1, 4).Value = _
        Array("Date", "Category", "Metric", "Value")

    Dim rank As Long, r As Long
    Dim used() As Boolean
    ReDim used(2 To lastRow)

    For rank = 1 To 5
        Dim bestIdx As Long, bestVal As Double
        bestIdx = -1: bestVal = -1E+18
        For r = 2 To lastRow
            If Not used(r) Then
                If dataWs.Cells(r, "D").Value > bestVal Then
                    bestVal = dataWs.Cells(r, "D").Value
                    bestIdx = r
                End If
            End If
        Next r
        If bestIdx = -1 Then Exit For
        used(bestIdx) = True
        dashWs.Cells(tableRow + 1 + rank, 2).Value = dataWs.Cells(bestIdx, "A").Value
        dashWs.Cells(tableRow + 1 + rank, 3).Value = dataWs.Cells(bestIdx, "B").Value
        dashWs.Cells(tableRow + 1 + rank, 4).Value = dataWs.Cells(bestIdx, "C").Value
        dashWs.Cells(tableRow + 1 + rank, 5).Value = bestVal
    Next rank

    Dim endRow As Long
    endRow = tableRow + 1 + rank - 1
    With dashWs.Range("B" & tableRow + 1 & ":E" & endRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    With dashWs.Range("B" & tableRow + 1 & ":E" & tableRow + 1)
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    dashWs.Range("B" & tableRow + 2 & ":B" & endRow).NumberFormat = "dd-mmm-yyyy"
    dashWs.Range("E" & tableRow + 2 & ":E" & endRow).NumberFormat = "#,##0"

    ' Conditional formatting: colour-scale the Value column of Top-5
    With dashWs.Range("E" & tableRow + 2 & ":E" & endRow).FormatConditions _
            .AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).FormatColor.Color = RGB(245, 130, 90)
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
    End With

    dashWs.Columns("B:E").AutoFit

    '----- Charts ------------------------------------------------------
    BuildCategoryChart dashWs, dataWs, lastRow, tableRow, endRow
    BuildTrendChart dashWs, dataWs, lastRow, tableRow, endRow

    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed from " & countRows & " data rows.", vbInformation, "Done"
End Sub

' Writes one KPI tile: label above, value below, coloured header band.
Private Sub WriteKpi(ByVal ws As Worksheet, ByVal topLeft As String, _
                     ByVal label As String, ByVal value As Variant, _
                     ByVal bandColor As Long)
    With ws.Range(topLeft).Resize(3, 2)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
    End With
    ws.Range(topLeft).Resize(1, 2).Interior.Color = bandColor
    ws.Range(topLeft).Value = label
    ws.Range(topLeft).Font.Color = RGB(255, 255, 255)
    ws.Range(topLeft).Font.Bold = True
    ws.Range(topLeft).Offset(1, 0).Value = value
    ws.Range(topLeft).Offset(1, 0).Font.Size = 18
    ws.Range(topLeft).Offset(1, 0).Font.Bold = True
    If IsNumeric(value) And value <> Int(value) Then
        ws.Range(topLeft).Offset(1, 0).NumberFormat = "#,##0.00"
    ElseIf IsNumeric(value) Then
        ws.Range(topLeft).Offset(1, 0).NumberFormat = "#,##0"
    End If
    ws.Range(topLeft).Offset(2, 0).Value = " "
End Sub

' Column chart: total Value per Category from the Data sheet.
Private Sub BuildCategoryChart(ByVal dashWs As Worksheet, _
                              ByVal dataWs As Worksheet, _
                              ByVal lastRow As Long, _
                              ByVal tableRow As Long, _
                              ByVal endRow As Long)
    Dim chObj As ChartObject
    Set chObj = dashWs.ChartObjects.Add( _
        Left:=dashWs.Range("B" & endRow + 3).Left, _
        Top:=dashWs.Range("B" & endRow + 3).Top, _
        Width:=340, Height:=220)
    With chObj.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Value by Category"
        .SetSourceData Source:=dataWs.Range("B1:B" & lastRow & ",D1:D" & lastRow)
        .SeriesCollection(1).XValues = dataWs.Range("B2:B" & lastRow)
        .HasLegend = False
    End With
End Sub

' Line chart: Value trend over Date from the Data sheet.
Private Sub BuildTrendChart(ByVal dashWs As Worksheet, _
                           ByVal dataWs As Worksheet, _
                           ByVal lastRow As Long, _
                           ByVal tableRow As Long, _
                           ByVal endRow As Long)
    Dim chObj As ChartObject
    Set chObj = dashWs.ChartObjects.Add( _
        Left:=dashWs.Range("G" & endRow + 3).Left, _
        Top:=dashWs.Range("G" & endRow + 3).Top, _
        Width:=340, Height:=220)
    With chObj.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Value Trend"
        .SetSourceData Source:=dataWs.Range("D1:D" & lastRow)
        .SeriesCollection(1).XValues = dataWs.Range("A2:A" & lastRow)
        .HasLegend = False
    End With
End Sub
