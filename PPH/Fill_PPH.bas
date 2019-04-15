Option Explicit

Sub InitTargetSheet(sht As Worksheet)
    ' PART 1 Delete All Cells
    sht.Cells.Select
    Selection.ClearContents
    Selection.NumberFormat = "General"

    ' PART 2 Delete Shapes Except Button
    Dim shp As Variant
    For Each shp In ActiveSheet.Shapes
        ' button.type = 12
        If shp.Type <> 12 Then
            shp.Delete
        End If
    Next

End Sub

Sub DrawCharts(chartsDict As Object)
    Dim aKeys As Variant: aKeys = chartsDict.keys
    Dim chartInx As Integer
    ' PART 1 Draw & Position Charts
    For chartInx = 0 To UBound(aKeys)
        Dim aItem As Variant: aItem = chartsDict(chartInx)
        Dim sRange As String
        ' --- Select Range
        sRange = "B3:C3," & aItem(0) & aItem(1) & ":" & aItem(2) & aItem(3)
        Range(sRange).Select ' Range("B3:C3,B4:C27").Select
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.ChartType = xlLine

        ' --- Set Data
        sRange = "PPH!$B$3:$C$3,PPH!" & "$" & aItem(0) & "$" & aItem(1) & ":" & "$" & aItem(2) & "$" & aItem(3) ' "PPH!$B$3:$C$3,PPH!$B$4:$C$27"
        ActiveChart.SetSourceData Source:=Range(sRange)

        ' --- Position
        sRange = "E" & aItem(1)
        Dim chartName As String ' "Chart 12345"
        Dim left As Integer: left = Range(sRange).left
        Dim top As Integer: top = Range(sRange).top
        chartName = "Chart " & Split(ActiveChart.name, " ")(2) ' ActiveChart.name = "PPH 图表 12345"
        ActiveSheet.Shapes(chartName).left = left
        ActiveSheet.Shapes(chartName).top = top

        ' --- Set Style
        ActiveChart.Axes(xlCategory).Select
        ActiveChart.Axes(xlCategory).TickLabelSpacing = 1
    Next

    ' END

End Sub

Function ParseCsvAndFillCell(resCsv As Variant)
    ' PART 1 Clean old Datas
    Application.ScreenUpdating = False
    Dim curSht As Worksheet: Set curSht = ActiveSheet
    Call InitTargetSheet(curSht)

    ' PART 2 Fill Cells By Two Dimentions Array
    Dim resCsvRowx As Integer: resCsvRowx = 0
    Dim sCurLine As String: sCurLine = ""
    Dim aRowData As Variant
    Dim head As String, tail As String, fmt As String
    Dim a2D(751, 3) As Variant ' ZhouYongHan : Max PPH Rows = 751, Max PPH Colx = 3
    Dim resCsvRows As Integer, resCsvCols As Integer: resCsvRows = 0: resCsvCols = 0

    ' PART 3 1st Read, Get Resource Csv's Rows & Cols
    Open resCsv For Input As #66
    Do While Not EOF(66)
        Line Input #66, sCurLine

        If resCsvRows = 0 Then
            aRowData = Split(sCurLine, ",")
            resCsvCols = UBound(aRowData)
        End If
        resCsvRows = resCsvRows + 1
    Loop
    Close #66

    ' PART 4 2rd Read, Fill a2D According Csv File & Some Global Dictionary
    Open resCsv For Input As #66
    Const meanRowx As Integer = 1
    Const dataColx As Integer = 1
    Dim rowx As Integer, serial As Integer: serial = 1
    Dim rangeDict As Object: Set rangeDict = CreateObject("Scripting.Dictionary") ' { 0:['B', 29, 'C', 52]}
    Dim rangeInx As Integer: rangeInx = 0
    For rowx = 1 To resCsvRows
        Line Input #66, sCurLine
        ' Fill With a Empty String
        If sCurLine = "" Then
            a2D(rowx, 0) = ""
        Else
            aRowData = Split(sCurLine, ",")
            Dim xInx As Integer: xInx = rowx - 1
            Dim yInx As Integer, cellValue As Variant
            For yInx = 0 To resCsvCols
                cellValue = aRowData(yInx)

                ' Get Mean
                If rowx = meanRowx Then cellValue = g_meanDict(cellValue)

                ' --- Set Serial
                If (yInx = dataColx) And (InStr(cellValue, " 01") > 0) Then
                    a2D(xInx, 0) = serial
                    serial = serial + 1
                    ' -- Set RangeDict
                    Dim aRange As Variant: aRange = Array("B", (rowx + 2), "C", (rowx + 2 + 23))  ' [B, 29, C, 52] ==> "B29", "C52"
                    rangeDict(rangeInx) = aRange
                    rangeInx = rangeInx + 1
                End If

                a2D(xInx, yInx) = cellValue
            Next
        End If
    Next
    Close #66

    ' ' PART 5 Fill Cells Start From A3 Cell
    Range("A3").Resize(resCsvRows + 1, resCsvCols + 1) = a2D

    Call DrawCharts(rangeDict)

    ' ' END
    ' Sheets(2).Activate
    Application.ScreenUpdating = True
    MsgBox "Success!"

End Function


