Option Explicit
' clear current sheet's cells
Function ClearCurSheet()
    Cells.Select
    Selection.ClearContents
End Function

Function ParseCsvAndFillCell(resCsv As Variant)
    ' START
    Application.ScreenUpdating = False

    ' PART 1
    Call ClearCurSheet

    ' PART 2
    call CreateSheets(g_sheetDict)

    ' PART 3
    Call FillSheetCells(resCsv)

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

    Const valueColx as Integer = 2
    Const cnColx as Integer = 3
    Const enColx as Integer = 4

    ' PART 2 Iterate csv file and fill Cells
    Do While Not EOF(1)
        Line Input #1, sCurLine
        aCsvRowData = Split(sCurLine, ",")

        Dim colx As Integer, fillColx as Integer, cellValue As String, DataID As String, group As String
        DataID = aCsvRowData(0) : fillColx = 0 :
        ' the top 3 lines's content is MoldHeader, 4th lines is unValid, others likes [ "0x400", 123, "我是中文翻译", "English Translation" ]
        If nCsvCurRowx > 4 Then
            Dim fillSheet As Worksheet, fillRowx as Integer
            ' create vaild groups only
            If g_groupDict.exists(DataID) Then
                group = g_groupDict(DataID)
                Set fillSheet = Sheets(group)
                ' fillRowx = fillSheet.Range("A65536").End(xlUp).Row + 1
                fillRowx = Application.CountA(fillSheet.Range("A:A")) + 1

                ' --- 遍历每行数据
                For colx = 0 To UBound(aCsvRowData)
                    Dim fmt As String : fmt = "General"
                    cellValue = aCsvRowData(colx)
                    fillColx = colx + 1

                    ' this colx need cell prec-format
                    If fillColx = valueColx then
                        Dim prec As Integer, head As String, tail As String
                        If g_precDict.exists(DataID) Then
                            prec = g_precDict(DataID)

                            Dim maxBitWeight As Variant, digit As Integer
                            maxBitWeight = 1
                            ' prec <> 0 Float Only
                            if prec <> 0 then
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
                            Else
                            ' prec = 0 : Positive Integer Only
                                fmt = "0"
                            End if

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
                    End if

                    If cellValue <> "" Then
                        ' fill each cell
                        With fillSheet.Cells(fillRowx, fillColx)
                            .NumberFormat = fmt
                            .FormulaR1C1 = cellValue
                        End With
                    End If
                Next
            End if
        Else
            Set fillSheet = Sheets(2)
            if nCsvCurRowx = 1 Then ' MoldName, SaveDate, Materials, Colour, MoldNum
                fillSheet.Range("A3").Resize(1, UBound(aCsvRowData) + 1) = aCsvRowData
            Elseif nCsvCurRowx = 2 Then ' 9, 2019/3/4-16:17:43, 1, 2, 3
                fillSheet.Range("A4").Resize(1, UBound(aCsvRowData) + 1) = aCsvRowData
            End if
        End if

        nCsvCurRowx = nCsvCurRowx + 1
    Loop
    Close #1
End Sub

Sub CreateSheets(sheetsDict As Object)
    Dim moldHeadSheet As Worksheet
    Set moldHeadSheet = Sheets(2)

    Call DelGroupSheets
    Dim aKeys As Variant, nInx As Integer
    aKeys = sheetsDict.keys

    For nInx = 0 To UBound(aKeys)
        Sheets.Add After:=moldHeadSheet
        ActiveSheet.Name = aKeys(nInx)
        Dim aTitle as Variant : aTitle = Array("DataID", "DataValue", "Description#1", "Description#2")
        ActiveSheet.Range("A1").Resize(1, UBound(aTitle) + 1) = aTitle
    Next
End Sub

Sub BeautySheets(sheetsDict As Object)
    Dim group as Variant
    For Each group in sheetsDict
        Dim sh as Worksheet : set sh = Sheets(group)

        ' Freeze The 1st Row
        sh.Activate
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
    For nInx = 1 To Sheets.Count
        If nInx > 2 Then
            ' the top two sheets is standard, delete others sheets only
            Worksheets(Sheets(3)).Delete
        End If
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
