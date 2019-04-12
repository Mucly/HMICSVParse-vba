Option Explicit

Sub InitTargetSheet(sht As Worksheet)
    ' PART 1 Delete All Cells
    sht.Activate
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

Function ParseCsvAndFillCell(resCsv As Variant)
    ' START
    Application.ScreenUpdating = False

    ' PART 1
    Dim parseSht As Worksheet : Set parseSht = Sheets(2)
    Call InitTargetSheet(parseSht)

    ' PART 2
    Call CreateSheets(g_sheetDict)

    ' PART 3
    Call FillSheetCells(resCsv)

    ' PART 4
    Call BeautySheets(g_sheetDict)

    ' END
    Sheets(2).Activate

    Application.ScreenUpdating = True  ' Restore
    MsgBox "Success！"

End Function
' TODO
Function GetArrayVaildCnt(a2D As Variant)
    Dim rowx As Integer, colx As Integer, nFillRowx As Integer
    nFillRowx = 0
    colx = 0
    For rowx = 0 To UBound(a2D)
        If a2D(rowx, 0) <> "" Then
            nFillRowx = nFillRowx + 1
        End If
    Next
    GetArrayVaildCnt = nFillRowx
End Function

Sub FillSheetCells(resCsv As Variant)
    ' PART 1 Read Csv By Line, Then format all cells in its sheets
    Dim sCurLine As String
    Dim aCsvRowData As Variant
    Dim nCsvCurRowx As Integer
    Open resCsv For Input As #1 ' csv.fileNumber == #1
    nCsvCurRowx = 1

    Const valueColx As Integer = 2
    Const cnColx As Integer = 3
    Const enColx As Integer = 4

    ' PART 2 Create a Sheet Named undefined
    g_sheetDict("undefined") = "undefined"  ' This sheet, called undefined, is used to save data that is not defined in the database sheet
    Dim undefinedSht As Worksheet : Set undefinedSht = Sheets("undefined")

    ' PART 3 Iterate csv file and fill Cells
    Do While Not EOF(1)
        Line Input #1, sCurLine
        if sCurLine <> "" Then
            aCsvRowData = Split(sCurLine, ",")

            Dim colx As Integer, cellValue As String, group As String
            Dim DataID As String : DataID = aCsvRowData(0)
            Dim fillColx As Integer : fillColx = 0
            ' the top 2 lines is MoldHeader, 3th lines is unValid, others likes [ "0x400", 123, "我是中文翻译", "English Translation" ]
            If nCsvCurRowx > 3 Then
                Dim fillSheet As Worksheet, fillRowx As Integer

                If g_groupDict.exists(DataID) Then
                    group = g_groupDict(DataID)
                    Set fillSheet = Sheets(group)
                Else
                    Set fillSheet = undefinedSht
                End if

                ' fillRowx = fillSheet.Range("A65536").End(xlUp).Row + 1
                fillRowx = Application.CountA(fillSheet.Range("A:A")) + 1

                ' --- 遍历每行数据
                For colx = 0 To UBound(aCsvRowData)
                    Dim fmt As String: fmt = "General"
                    cellValue = aCsvRowData(colx)
                    fillColx = colx + 1

                    ' this colx need cell prec-format
                    If fillColx = valueColx Then
                        Dim prec As Integer, head As String, tail As String
                        If g_precDict.exists(DataID) Then
                            prec = g_precDict(DataID)

                            Dim maxBitWeight As Variant, digit As Integer
                            maxBitWeight = 1
                            ' prec <> 0 Float Only
                            If prec <> 0 Then
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
                                    fmt = head + "." + tail
                                End If
                            Else
                            ' prec = 0 : Positive Integer Only
                                fmt = "0"
                            End If

                            cellValue = Format(cellValue / maxBitWeight, fmt)
                        End If
                    ' this colx need get cn trans
                    ElseIf fillColx = cnColx Then
                        If g_cnDict.exists(DataID) Then
                            cellValue = g_cnDict(DataID)
                            ' Debug.print "group=", group, "   DataID=", DataID, " CN=", cellValue
                        End If
                    ' this colx need get en trans
                    ElseIf fillColx = enColx Then
                        If g_enDict.exists(DataID) Then
                            cellValue = g_enDict(DataID)
                            ' Debug.print "group=", group, "   DataID=", DataID, " EN=", cellValue
                        End If
                    ' default colx
                    Else
                        cellValue = cellValue
                    End If

                    If cellValue <> "" Then
                        ' fill each cell
                        With fillSheet.Cells(fillRowx, fillColx)
                            .NumberFormat = fmt
                            .FormulaR1C1 = cellValue
                        End With
                    End If
                Next
            Else
                Set fillSheet = Sheets(2)
                If nCsvCurRowx = 1 Then ' MoldName, SaveDate, Materials, Colour, MoldNum
                    fillSheet.Range("A3").Resize(1, UBound(aCsvRowData) + 1) = aCsvRowData
                ElseIf nCsvCurRowx = 2 Then ' 9, 2019/3/4-16:17:43, 1, 2, 3
                    fillSheet.Range("A4").Resize(1, UBound(aCsvRowData) + 1) = aCsvRowData
                End If
            End If

            nCsvCurRowx = nCsvCurRowx + 1
        End if
    Loop
    Close #1

    ' END hidden unvaild sheet
    undefinedSht.Visible = False

End Sub

Sub CreateSheets(sheetsDict As Object)
    Dim moldHeadSheet As Worksheet
    Set moldHeadSheet = Sheets(2)

    Call DelGroupSheets
    Dim aKeys As Variant, nInx As Integer
    aKeys = sheetsDict.keys

    Dim newShtCnts As Integer: newShtCnts = UBound(aKeys) + 1
    For nInx = 0 To newShtCnts
        Sheets.Add After:=moldHeadSheet

        If nInx <> newShtCnts Then
            ActiveSheet.Name = aKeys(nInx)

            Dim aTitle As Variant: aTitle = Array("DataID", "DataValue", "Description#1", "Description#2")
            ActiveSheet.Range("A1").Resize(1, UBound(aTitle) + 1) = aTitle
        Else
            ActiveSheet.Name = "Merge"
            ActiveWindow.SelectedSheets.Visible = False
        End If

    Next
End Sub

Sub BeautySheets(sheetsDict As Object)
    Dim group As Variant
    For Each group In sheetsDict
        Dim sht As Worksheet: Set sht = Sheets(group)

        ' Freeze The 1st Row
        sht.Activate
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True

        ' Cells Format
        With Cells
            .Columns.AutoFit
            .HorizontalAlignment = xlHAlignCenter
            .Font.Name = "微软雅黑"
            .Font.Size = 12
        End With
    Next
End Sub

Sub DelGroupSheets()
    Application.DisplayAlerts = False
    Dim nInx As Integer
    ' sheet's index start from 1
    For nInx = 3 To Sheets.Count
        ' the top two sheets is standard, delete others sheets only
        Sheets(3).Delete
    Next
    Application.DisplayAlerts = True
End Sub

' sheet Protect
Sub LockMoldHeader(bLocked)
    If (bLocked) Then
        Cells.Select
        Selection.Locked = False
        Range("A1:K6").Locked = bLocked
        ActiveSheet.Protect "dfg312"
    Else
        ActiveSheet.Unprotect
    End If

End Sub
