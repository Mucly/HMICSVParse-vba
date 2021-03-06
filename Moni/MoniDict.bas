Option Explicit
Public g_meanDict, g_precDict, g_visbDict As Object ' 多国语言字典
' * 获取多国语言翻译字典  k=DataID，v=translated content
' * ！ 行列从1开始计数
' @ Parameter  DB - string - 目标DB sheet的名字
Function GetTransDict(DB As String)
    ' PART 1 创建DB worksheet对象
    Dim DBWorkSheet As Worksheet
    Set DBWorkSheet = Worksheets(DB)

    ' PART 2 创建字典
    Set g_meanDict = CreateObject("Scripting.Dictionary")
    Set g_precDict = CreateObject("Scripting.Dictionary")
    Set g_visbDict = CreateObject("Scripting.Dictionary")

    ' PART 3 遍历DB worksheet，并建立字典映射关系
    Const tagColx  As Integer = 1
    Const meanColx  As Integer = 2
    Const precColx  As Integer = 3
    Const visbColx As Integer = 4

    Dim key As String, mean As String, prec As String, visb As String
    Dim rowx As Integer

    For rowx = 2 To DBWorkSheet.UsedRange.Rows.Count ' i表示行号，第一行不需要加入字典，故从2开始
        key = DBWorkSheet.Cells(rowx, tagColx).Value
        If key <> "" Then
            mean = DBWorkSheet.Cells(rowx, meanColx)
            prec = DBWorkSheet.Cells(rowx, precColx)
            visb = DBWorkSheet.Cells(rowx, visbColx)

            g_meanDict(key) = mean
            g_precDict(key) = prec
            g_visbDict(key) = visb
        End If
    Next
End Function


