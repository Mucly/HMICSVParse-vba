Option Explicit

Sub ClearCurSheet()
    Cells.Select
    Selection.ClearContents
    Selection.NumberFormat = "General"
End Sub

Sub BeautySheets()
    ' PART 1 Format Time Colx
    Const dataColx As Integer = 2
    Dim maxCols  As Integer: maxCols = Application.CountA(ActiveSheet.Range("A:A")) + 3
    Range("B3:B" & maxCols).NumberFormat = "yyyy-m-d hh:mm:ss"
End Sub

Sub FormatColWithPrecison(precDict As Object)
    Dim colx As Integer
    Dim aKeys As Variant: aKeys = precDict.keys
    Dim keysCnt As Integer: keysCnt = UBound(aKeys)
    For colx = 0 To keysCnt
        Dim prec As Integer: prec = precDict(colx)
        If prec <> 0 Then
            Dim fmt As String
            fmt = "0." & String(prec, "0") & "_ "
            Range(g_colxAlphaDict(colx) & ":" & g_colxAlphaDict(colx)).NumberFormatLocal = fmt
        End If
    Next
End Sub

Sub DelEachSegSheets()
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

Sub CreateEachSegSheets()
    ' PART 1 Del Sheets
    Call DelEachSegSheets

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

Function ParseCsvAndFillCell(resCsv As Variant, fillRowx As Integer)
    ' PART 1 Clean old Datas
    Application.ScreenUpdating = False
    Call ClearCurSheet

    ' PART 2 Fill Cells By Lines, Not By Two Dimentions Array
    Dim resCsvRowx As Integer: resCsvRowx = 0
    Dim sCurLine As String: sCurLine = ""
    Dim aRowData As Variant
    Dim head As String, tail As String, fmt As String
    Dim a2D(7300, 120) As Variant ' Max Temper Rows < 7300, Max Temper Colx < 120
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

        Dim resCsvColx As Integer, cellValue As Variant ' resCsvColx != colx, resCsvColx START FROM 0
        ' Set Data's Precsion Which Belong to Current Row Array
        For resCsvColx = 0 To resCsvCols
            fmt = "General"

            Dim fillColx As Integer: fillColx = resCsvColx + 1
            cellValue = aRowData(resCsvColx)
            ' Translate Title(Rowx1) And Reset g_tagPrecDict
            If resCsvRowx = 0 Then
                ' 1 - Get Translate
                If g_meanDict.exists(cellValue) Then
                    cellValue = g_meanDict(cellValue)
                End If
                ' 2 - Get Precison Colx
                If (resCsvColx >= 2) And g_tagPrecDict.exists(aRowData(resCsvColx)) Then
                    g_colxPrecDict(fillColx) = g_tagPrecDict(aRowData(resCsvColx))
                End If
            Else
            ' Those Rowx Need Precsion
            ' Note : The Maximum Precsion is ONE !
                If g_colxPrecDict(fillColx) <> 0 Then
                    Dim prec As Integer
                    If g_colxPrecDict.exists(fillColx) Then
                        prec = g_colxPrecDict(fillColx)
                        ' prec <> 0 Float Only
                        If prec <> 0 Then
                            Dim maxBitWeight As Variant, digit As Integer
                            maxBitWeight = 1

                            digit = Len(cellValue)
                            maxBitWeight = Application.WorksheetFunction.Power(10, prec)
                            If prec > digit Then
                                head = "0"
                                tail = String(prec, "0")
                                fmt = head + "." + tail
                            ElseIf prec < digit Then
                                head = "0"
                                tail = String(prec, "0")
                                fmt = head + "." + tail
                            Else
                                head = "0"
                                tail = String(prec, "0")
                                fmt = head + "." + tail
                            End If
                            cellValue = Format(cellValue / maxBitWeight, fmt)
                        End If
                    End If
                End If
            End If

            a2D(resCsvRowx, resCsvColx) = cellValue
        Next

        ' Counter Accumulate
        resCsvRowx = resCsvRowx + 1
        fillRowx = fillRowx + 1
    Loop
    Close #66

    ' PART 5 Fill Cells Start From A3 Cell
    Range("A3").Resize(resCsvRows + 1, resCsvCols + 1) = a2D

    ' PART 5 Format Columns
    Call FormatColWithPrecison(g_colxPrecDict)

    ' PART 6 Beauty And Create Sheets, Then Draw Charts
    Call BeautySheets
    Call CreateEachSegSheets

    ' END
    Sheets(2).Activate
    Application.ScreenUpdating = True
    MsgBox "Success!"

End Function

