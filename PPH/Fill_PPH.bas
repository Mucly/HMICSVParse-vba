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

Sub BeautySheets()
    ' PART 1 Format Time Colx
    Const dataColx As Integer = 2
    Dim maxCols  As Integer: maxCols = Application.CountA(ActiveSheet.Range("A:A")) + 3
    Range("B3:B" & maxCols).NumberFormat = "yyyy-m-d hh:mm:ss"
End Sub

Sub DrawCharts()
    ' PART 2 Draw Charts
    Dim aSegSheetName As Variant: aSegSheetName = g_meanDict.Items
    Dim inx As Integer
    Dim parseSht As Worksheet: Set parseSht = Sheets(2)
    Dim name As Integer: name = 1
    Dim maxCols  As Integer: maxCols = Application.CountA(ActiveSheet.Range("A:A")) + 2
    For inx = 2 To UBound(aSegSheetName)
        Dim colx As Integer: colx = inx + 1
        ' Each Chart's Title Depend On Odd Colx's Title
        If (colx Mod 2) <> 0 Then
            Sheets.Add After:=parseSht
            ActiveSheet.name = "#" & name

            Dim sTimeRange As String, sTemperRange As String, sRange As String
            sTimeRange = "Temper!$B$3" & ":$B$" & maxCols ' Time Col， eg. "Temper!$B3:B$16"
            sTemperRange = "Temper!$" & g_colxAlphaDict(colx) & "$3" & ":$" & g_colxAlphaDict(colx + 1) & "$" & maxCols ' Temper Cols， eg. Temper!$D$3:$E$6"
            sRange = sTimeRange & "," & sTemperRange

            ActiveSheet.Shapes.AddChart.Select
            ActiveChart.ChartType = xlLine
            ' ActiveChart.SetSourceData Source:=Range("Temper!$B$3:$B$6,Temper!$D$3:$E$6")
            ActiveChart.SetSourceData Source:=Range(sRange)
            ActiveChart.ApplyLayout (3)
            ActiveChart.Axes(xlCategory).Select
            ActiveChart.Axes(xlCategory).CategoryType = xlCategoryScale
            ActiveChart.ChartTitle.Select
            ActiveChart.ChartTitle.Text = "Temper #" & name
            Selection.Format.TextFrame2.TextRange.Characters.Text = "Temper #" & name

            ' Cells(6,1).Select
            name = name + 1
        End If
    Next

End Sub

Function ParseCsvAndFillCell(resCsv As Variant)
    ' PART 1 Clean old Datas
    Application.ScreenUpdating = False
    Dim curSht As Worksheet : Set curSht = ActiveSheet
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
    Dim rowx As Integer, serial As Integer : serial = 1
    For rowx = 1 To resCsvRows
        Line Input #66, sCurLine
        if sCurLine = "" Then
            a2D(rowx, 0) = ","
        Else
            aRowData = Split(sCurLine, ",")
            Dim xInx As Integer : xInx = rowx - 1
            Dim yInx As Integer, cellValue As Variant
            For yInx = 0 To resCsvCols
                cellValue = aRowData(yInx)

                ' Get Mean
                if rowx = meanRowx Then cellValue = g_meanDict(cellValue)

                ' --- Set Serial
                if (yInx = dataColx) And (instr(cellValue, " 01") > 0) Then
                    a2D(xInx, 0) = serial
                    serial = serial + 1
                End If

                a2D(xInx, yInx) = cellValue
            Next
        End if
    Next
    Close #66

    ' ' PART 5 Fill Cells Start From A3 Cell
    Range("A3").Resize(resCsvRows + 1, resCsvCols + 1) = a2D

    Call DrawCharts

    ' ' END
    ' Sheets(2).Activate
    Application.ScreenUpdating = True
    MsgBox "Success!"

End Function

