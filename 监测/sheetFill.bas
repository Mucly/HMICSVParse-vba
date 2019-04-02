Option Explicit
' * 每次载入前，清空所有单元格 ---
Function ClearCurSheet()
    Cells.Select
    Selection.ClearContents
End Function

' * 读取csv文件并对照字典翻译后，填充至当前sheet
' @Parameter  resCsv        - 源csv文件
'                      offsetRowx - 填充行偏移位置
' @Remind sheet里面的行列从1开始
Function ParseCsvAndFillCell(resCsv As Variant, offsetRowx As Integer)
    ' 禁用自动刷新，计算模式为手动
    Application.ScreenUpdating = False

    ' PART 1 清空当前sheet的所有单元格内容后，关闭屏幕更新，提升解析速度
    ClearCurSheet

    ' PART 2 设置数据对象
    Dim fillRowx As Integer
    fillRowx = offsetRowx ' fill the cell start from the third row

    ' PART 3 逐行读取csv文件
    Dim curLine As String, fillContent As String ' 当前行内容（字符串格式）、填充的内容
    ' Range("A3").Resize(rows, colx).Value = a2D

    ' PART 2 第一次遍历，构造一个二维数据，为了加快单元格填充速度
    Dim n2DRows As Integer ' 文件总行数
    Dim n2DCols As Integer
    Open resCsv For Input As #1 ' 打开csv文件，file number #1
    Do While Not EOF(1) ' 逐行循环 #1文件
        Line Input #1, curLine

        If n2DCols = 0 Then
            Dim aRowData As Variant
            aRowData = Split(curLine, ",")
            n2DCols = UBound(aRowData)
        End If

        n2DRows = n2DRows + 1

    Loop
    Close #1

    ' PART 3 第二次遍历，给二维数组赋值
    Dim csvCurRowx As Integer
    Dim a2D(0 To 5001, 0 To 50) As Variant ' 这里只能常量
    csvCurRowx = 0
    Open resCsv For Input As #1
    Do While Not EOF(1) ' 逐行循环 #1文件
        Line Input #1, curLine

        Dim inx As Integer, inx2 As Integer
        inx2 = 0
        For inx = 0 To n2DCols
            Dim aRowData2 As Variant
            Dim key as String
            key = inx & ""

            if csvCurRowx = 0 Then
                aRowData2 = Split(curLine, ",")
                if g_meanDict.exists(key) Then
                    aRowData2(inx) = g_meanDict(key)
                End if
            Else ' 补上序号
                curLine = csvCurRowx & curLine
                aRowData2 = Split(curLine, ",")

                Dim prec as Integer
                if g_precDict.exists(key) Then
                    prec = Val(g_precDict(key))
                Else
                    prec = 0
                End if

                If prec <> 0 Then
                    Dim maxBitWight as Integer, digit as Integer
                    maxBitWight = Application.WorksheetFunction.Power(10, prec)
                    digit = Len(aRowData2(inx)) ' 源数字的位数

                    Dim head As String, tail As String, fmt As String   ' 配置format的格式
                    fmt = "General"

                    ' #4 精度 > 位数
                    If prec > digit Then
                        head = "0"
                        tail = String(prec, "0")
                        fmt = head + "." + tail
                    ' 精度 < 位数
                    ElseIf prec < digit Then
                        head = String(digit - prec, "0")
                        tail = String(prec, "0")
                        head = "0"
                        fmt = head + "." + tail
                    ' 精度 = 位数
                    Else
                        head = "0"
                        tail = String(prec, "0")
                        fmt = head + "." + tail
                    End If   ' #4
                    aRowData2(inx) = Format(aRowData2(inx) / maxBitWight, fmt)
                End If
            End if

            if g_visbDict(key) <> 0 Then
                a2D(csvCurRowx, inx2) = aRowData2(inx)
                inx2 = inx2 + 1
            end if
        Next

        csvCurRowx = csvCurRowx + 1
    Loop
    Close #1

    Range("A3").Resize(n2DRows + 1, n2DCols + 1) = a2D

    ' END
    Application.ScreenUpdating = True  ' 还原屏幕刷新
    MsgBox "解析成功！" ' 弹出成功提示

End Function
