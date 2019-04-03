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
    Call ClearCurSheet()

    ' PART 2 第一次逐行读取csv文件，获取csv文件的总行、列数，用以绘制单元格内容
    Dim sCurLine As String
    Dim n2DRowsNum As Integer, n2DColsNum As Integer
    Dim aRowData As Variant

    Open resCsv For Input As #1 ' 打开csv文件，file number #1
    ' --- 逐行读取csv文件，获取csv文件的总行、列数
    Do While Not EOF(1)
        Line Input #1, sCurLine

        ' 通过csv的第一行内容
        If n2DColsNum = 0 Then
            aRowData = Split(sCurLine, ",")
            n2DColsNum = UBound(aRowData)
        End If
        n2DRowsNum = n2DRowsNum + 1
    Loop
    Close #1

    ' PART 3 第二次逐行读取csv文件，给二维数组赋值
    const nMeanColx as Integer = 4
    Dim csvCurRowx As Integer, a2D(0 To 5000, 0 To 5) As Variant ' 这里只能常量
    csvCurRowx = 0
    Open resCsv For Input As #1
    Do While Not EOF(1) ' 逐行循环 #1文件
        Line Input #1, sCurLine

        ' csv其他行，需要在前面加上编号
        if csvCurRowx <> 0 Then
            sCurLine = csvCurRowx & sCurLine
        End If
        aRowData = Split(sCurLine, ",")

        Dim k as String, v as String, nColx as Integer
        For nColx = 0 To UBound(aRowData)
            k = aRowData(nColx)
            ' csv第一行，直接进行翻译
            if csvCurRowx = 0 Then
                IF g_meanDict.exists(k) Then
                    v = g_meanDict(k)
                Else
                    v = k
                End If
            Else
                ' 第四列需要进行字典翻译
                if (nColx + 1) = nMeanColx Then
                    IF g_meanDict.exists(k) Then
                        v = g_meanDict(k)
                    Else
                        v = k
                    End If
                Else
                    v = k
                End if
            End if

            a2D(csvCurRowx, nColx) = v
        Next

        csvCurRowx = csvCurRowx + 1
    Loop
    Close #1

    ' PART 4 以A3单元格为原点，按照二维数组数据进行单元格批量绘制
    Range("A3").Resize(n2DRowsNum + 1, n2DColsNum + 1) = a2D

    ' END
    Application.ScreenUpdating = True  ' 还原屏幕刷新
    MsgBox "解析成功！" ' 弹出成功提示

End Function

