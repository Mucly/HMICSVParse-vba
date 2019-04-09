Option Explicit

Sub ClearCurSheet()
    Cells.Select
    Selection.ClearContents
    Selection.NumberFormat = "General"
End Sub

Sub BeautySheets()

End Sub

Sub ShowCharts()
    ActiveSheet.Shapes.AddChart.Select

    ActiveChart.ChartType = xlLineMarkersStacked
    ActiveChart.SetSourceData Source:=Range("Temper!$A$3:$T$1299")
    ActiveChart.SeriesCollection(1).Delete
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).Name = "=Temper!$C$3"
    ActiveChart.SeriesCollection(1).Values = "=Temper!$C$4:$C$6"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=Temper!$D$3"
    ActiveChart.SeriesCollection(2).Values = "=Temper!$D$4:$D$6"
    ActiveChart.SeriesCollection(2).XValues = "=Temper!$A$4:$B$6"
End Sub

Function ParseCsvAndFillCell(resCSV As Variant, fillRowx As Integer)
    ' PART 1 Clean old Datas
    Application.ScreenUpdating = False
    call ClearCurSheet

    ' PART 2 Fill Cells By Lines, Not By Two Dimentions Array
    Dim resCSVRowx As Integer : resCSVRowx = 0
    Dim sCurLine As String : sCurLine = ""
    Dim aRowData As Variant
    Dim head As String, tail As String, fmt As String

    Open resCSV For Input As #66
    Do While Not EOF(66)
        Line Input #66, sCurLine
        aRowData = Split(sCurLine, ",")

        Dim colInx As Integer, cellValue As Variant
        Dim resCols As Integer : resCols = UBound(aRowData)
        Dim resColsAdd1 As Integer : resColsAdd1 = resCols + 1 ' Colx Start From 1, So need Add 1

        ' Set Data's Precsion Which Belong to Current Row Array
        For colInx = 0 To resCols
            fmt = "General"

            Dim fillColx As Integer : fillColx = colInx + 1
            cellValue = aRowData(colInx)
            ' Translate Title(Rowx1) And Reset g_precDict
            if resCSVRowx = 0 Then
                ' 1 - Get Translate
                if g_meanDict.exists(cellValue) Then
                    aRowData(colInx) = g_meanDict(cellValue)
                End if
                ' 2 - Get Precison Colx
                If (colInx >= 2) AND g_precDict.exists(cellValue) Then
                    g_precDict(colInx) = g_precDict(cellValue)
                End if
            Else
            ' Those Rowx Need Precsion
            ' Note : The Maximum Precsion is ONE !
                if g_precDict(colInx) <> 0 Then
                    Dim prec As Integer
                    If g_precDict.exists(colInx) Then
                        prec = g_precDict(colInx)
                        ' prec <> 0 Float Only
                        if prec <> 0 then
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
                                fmt = head + "." +  tail
                            End If
                            cellValue = Format(cellValue / maxBitWeight, fmt)
                            ' fmt = fmt & "_ "
                        End if
                    End If
                End if
            End if

            Cells(fillRowx, fillColx).Value = cellValue
            ' TODO
            ' When Precsion > 1, Open Those Code
            ' Optimize : Gets the column number that needs precision and formats it once
            ' If fmt = "General" Then
            '     Cells(fillRowx, fillColx).Value = cellValue
            ' Else
            '     With Cells(fillRowx, fillColx)
            '         .Value = cellValue
            '         .NumberFormatLocal = fmt
            '     End With
            ' End If

        Next

        ' Counter Accumulate
        resCSVRowx = resCSVRowx + 1
        fillRowx = fillRowx + 1
    Loop
    Close #66


    ' END
    Application.ScreenUpdating = True
    MsgBox "Success!"

End Function
