Option Explicit

Sub ClearCurSheet()
    Cells.Select
    Selection.ClearContents
    ' Selection.NumberFormat = "General"
End Sub

Sub BeautySheets()
    ' PART 1 Format Time Colx
    ' Const dataColx As Integer = 2
    ' Dim maxCols  As Integer: maxCols = Application.CountA(ActiveSheet.Range("A:A")) + 3
    ' Range("B3:" & "B" & maxCols).NumberFormat = "yyyy-m-d h:mm:ss"

    ' PART 2 Format Each Precison Colx
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
    Call DelEachSegSheets()

    ' PART 2 Draw Charts
    Dim aSegSheetName As Variant: aSegSheetName = g_meanDict.Items
    Dim inx As Integer
    Dim parseSht As Worksheet: Set ParseSht = Sheets(2)
    Dim name As Integer : name = 1
    Dim maxCols  As Integer: maxCols = Application.CountA(ActiveSheet.Range("A:A")) + 2
    For inx = 2 To UBound(aSegSheetName)
        Dim colx As Integer : colx = inx + 1
        ' Each Chart's Title Depend On Odd Colx's Title
        If (colx Mod 2) <> 0 Then
            Sheets.Add After := ParseSht
            ActiveSheet.Name = "#" & name

            Dim sTimeRange As String, sTemperRange As String, sRange As String
            sTimeRange = "Temper!$B$3" & ":$B$" & maxCols ' Time Col， eg. "Temper!$B3:B$666"
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

Function ParseCsvAndFillCell(resCSV As Variant, fillRowx As Integer)
    ' PART 1 Clean old Datas
    Application.ScreenUpdating = False
    Call ClearCurSheet

    ' PART 2 Fill Cells By Lines, Not By Two Dimentions Array
    Dim resCSVRowx As Integer: resCSVRowx = 0
    Dim sCurLine As String: sCurLine = ""
    Dim aRowData As Variant
    Dim head As String, tail As String, fmt As String
    REM Dim dFormatColx As Object : Set dFormatColx = CreateObject("Scripting.Dictionary") '

    Open resCSV For Input As #66
    Do While Not EOF(66)
        Line Input #66, sCurLine
        aRowData = Split(sCurLine, ",")

        Dim colInx As Integer, cellValue As Variant ' colInx != colx, colInx START FROM 0
        Dim resCols As Integer: resCols = UBound(aRowData)
        Dim resColsAdd1 As Integer: resColsAdd1 = resCols + 1  ' Colx Start From 1, So need Add 1

        ' Set Data's Precsion Which Belong to Current Row Array
        For colInx = 0 To resCols
            fmt = "General"

            Dim fillColx As Integer: fillColx = colInx + 1
            cellValue = aRowData(colInx)
            ' Translate Title(Rowx1) And Reset g_precDict
            If resCSVRowx = 0 Then
                ' 1 - Get Translate
                If g_meanDict.exists(cellValue) Then
                    cellValue = g_meanDict(cellValue)
                End If
                ' 2 - Get Precison Colx
                If (colInx >= 2) And g_precDict.exists(aRowData(colInx)) Then
                    g_precDict(colInx) = g_precDict(aRowData(colInx))
                End If
            Else
            ' Those Rowx Need Precsion
            ' Note : The Maximum Precsion is ONE !
                If g_precDict(colInx) <> 0 Then
                    Dim prec As Integer
                    If g_precDict.exists(colInx) Then
                        prec = g_precDict(colInx)
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
                            fmt = fmt + "_ " ' Selection.NumberFormatLocal = "0.00_ "
                        End If
                    End If
                End If
            End If

            Cells(fillRowx, fillColx).Value = cellValue

            ' TODO
            ' When Precsion > 1, Open Those Code
            ' * Optimize : Gets the column number that needs precision and formats it once
            ' * cell with precsion   =>   fmt = fmt & "_ "
            If fmt = "General" Then
                Cells(fillRowx, fillColx).Value = cellValue
            Else
                With Cells(fillRowx, fillColx)
                    .Value = cellValue
                    .NumberFormatLocal = fmt
                End With
            End If

        Next

        ' Counter Accumulate
        resCSVRowx = resCSVRowx + 1
        fillRowx = fillRowx + 1
    Loop
    Close #66

    Call BeautySheets
    Call CreateEachSegSheets

    ' END
    Application.ScreenUpdating = True
    MsgBox "Success!"

End Function
