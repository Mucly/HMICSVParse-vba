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

    ' PART 2 第一次逐行读取csv文件，获取csv文件的总行、列数，用以绘制单元格内容
    Dim sCurLine As String
    Dim n2DRows As Integer, n2DCols As Integer
    Open resCsv For Input As #1 ' 打开csv文件，file number #1
    ' --- 逐行读取csv文件，获取csv文件的总行、列数
    Do While Not EOF(1)
        Line Input #1, sCurLine

        ' 通过csv的第一行内容
        If n2DCols = 0 Then
            Dim aRowData As Variant
            aRowData = Split(sCurLine, ",")
            n2DCols = UBound(aRowData)
        End If
        n2DRows = n2DRows + 1
    Loop
    Close #1

    ' PART 3 第二次逐行读取csv文件，给二维数组赋值
    Dim csvCurRowx As Integer
    Dim a2D(0 To 5001, 0 To 50) As Variant ' 这里只能常量
    csvCurRowx = 0
    Open resCsv For Input As #1
    Do While Not EOF(1) ' 逐行循环 #1文件
        Line Input #1, sCurLine

        ' 因为有需要隐藏的数据列，故建立一个inx2，隐藏的数据就不用写到二维数组内了
        Dim inx As Integer, inx2 As Integer
        inx2 = 0
        For inx = 0 To n2DCols
            Dim aRowData2 As Variant
            Dim key as String
            key = (inx - 1) & "" ' key = -1，代表第一列

            ' 将csv的第一行，按照meanDict进行翻译
            if csvCurRowx = 0 Then
                aRowData2 = Split(sCurLine, ",")
                if g_meanDict.exists(key) Then
                    aRowData2(inx) = g_meanDict(key)
                End if
            Else ' 第二行开始的数据，需在前面补上序号
                sCurLine = csvCurRowx & sCurLine
                aRowData2 = Split(sCurLine, ",")

                Dim prec as Integer
                if g_precDict.exists(key) Then
                    prec = g_precDict(key)
                Else
                    prec = 0
                End if

                ' 如果数据精度不为0，则进行精度格式化处理
                If prec <> 0 Then
                    Dim nMaxBitWeight as Integer, nDigit as Integer
                    nMaxBitWeight = Application.WorksheetFunction.Power(10, prec) ' 数字的权位（重）
                    nDigit = Len(aRowData2(inx)) ' 源数字的位数

                    Dim sHead As String, sTail As String, sFmt As String   ' 配置format的格式
                    sFmt = "General"

                    ' 精度 > 位数
                    If prec > nDigit Then
                        sHead = "0"
                        sTail = String(prec, "0")
                        sFmt = sHead + "." + sTail
                    ' 精度 < 位数
                    ElseIf prec < nDigit Then
                        sTail = String(prec, "0")
                        sHead = "0"
                        sFmt = sHead + "." + sTail
                    ' 精度 = 位数
                    Else
                        sHead = "0"
                        sTail = String(prec, "0")
                        sFmt = sHead + "." + sTail
                    End If
                    aRowData2(inx) = Format(aRowData2(inx) / nMaxBitWeight, sFmt)
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

    ' PART 4 以A3单元格为原点，按照二维数组数据进行单元格批量绘制
    Range("A3").Resize(n2DRows + 1, n2DCols + 1) = a2D

    ' END
    Application.ScreenUpdating = True  ' 还原屏幕刷新
    MsgBox "解析成功！" ' 弹出成功提示

End Function
