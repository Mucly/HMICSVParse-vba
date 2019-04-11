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

Sub DelAllCharts()
    Application.DisplayAlerts = False
    Dim nInx As Integer
    ' sheet's index start from 1
    For nInx = 1 To Sheets.Count
        If nInx > 2 Then
            ' the top two sheets is standard, delete others sheets only
            Sheets(3).Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

Sub DrawCharts()
    ' PART 1 Del Sheets
    Call DelAllCharts

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
    Do While Not EOF(66)
        Line Input #66, sCurLine
        aRowData = Split(sCurLine, ",")

        Dim resInx As Integer, cellValue As Variant
        For resInx = 0 To resCsvCols
            Dim fillColx As Integer: fillColx = resInx + 1
            a2D(resCsvRowx, resInx) = cellValue
        Next

        ' Counter Accumulate
        resCsvRowx = resCsvRowx + 1
    Loop
    Close #66

    ' ' PART 5 Fill Cells Start From A3 Cell
    Range("A3").Resize(resCsvRows + 1, resCsvCols + 1) = a2D

    ' Call DrawCharts

    ' ' END
    ' Sheets(2).Activate
    Application.ScreenUpdating = True
    MsgBox "Success!"

End Function

