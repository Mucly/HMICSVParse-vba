Option Explicit

Sub ClearCurSheet()
    Cells.Select
    Selection.ClearContents
End Sub

Sub BeautySheets()
End Sub

Sub ShowCharts()
End Sub

Function ParseCsvAndFillCell(resCSV As Variant, fillRowx As Integer)
    ' PART 1 Clean old Datas
    Application.ScreenUpdating = False
    call ClearCurSheet

    ' PART 2 Fill Cells By Lines, Not By Two Dimentions Array
    Dim resCSVRowx As Integer : resCSVRowx = 0
    Dim sCurLine As String : sCurLine = ""
    Dim aRowData As Variant
    Dim head As String, tail As String, fmt As String : fmt = "General"

    Open resCSV For Input As #66
    Do While Not EOF(66)
        Line Input #66, sCurLine
        aRowData = Split(sCurLine, ",")

        Dim colInx As Integer, cellValue As String
        Dim resCols As Integer : resCols = UBound(aRowData)
        Dim resColsAdd1 As Integer : resColsAdd1 = resCols + 1 ' Colx Start From 1, So need Add 1

        ' Set Data's Precsion Which Belong to Current Row Array
        For colInx = 0 To resCols
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
                    Debug.print colInx, g_precDict(colInx)
                End if
            Else
            ' Those Rowx Need Precsion
                if g_precDict(colInx) <> 0 Then
                    Dim key As String : key = g_meanDict(colInx)
                    Dim prec As Integer
                    If g_precDict.exists(colInx) Then
                        prec = g_precDict(colInx)
                        ' prec <> 0 Float Only
                        if prec <> 0 then
                            Dim maxBitWeight As Variant, digit As Integer
                            maxBitWeight = 1

                            cellValue = Replace(cellValue, ".", "")
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
                            aRowData(colInx) = Format(cellValue / maxBitWeight, fmt)
                        End if
                    End If
                End if
            End if
        Next

        ' Fill Cells
        Range("A" & fillRowx).Resize(1, resColsAdd1) = aRowData

        ' Counter Accumulate
        fillRowx = fillRowx + 1
        resCSVRowx = resCSVRowx + 1
    Loop
    Close #66


    ' END
    Application.ScreenUpdating = True
    MsgBox "Success!"

End Function
