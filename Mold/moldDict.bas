Option Explicit
Public g_groupDict As Object ' 资料分组信息 { dataID:'模具' }
Public g_datasDict As Object ' 填充数据分组信息 { 'sheet名字（这里是分组）' : 该sheet行列数据组成的单元格 }
' * 获取多国语言翻译字典  k=DataID，v=translated content
' @Parameter  DB - string - 目标DB sheet的名字
Function GetTransDict(DB As String)
    ' PART 1 创建DB worksheet对象
    Dim DBWorkSheet As Worksheet
    Set DBWorkSheet = Worksheets(DB)

    ' PART 2 创建多国语言字典
    Set g_groupDict = CreateObject("Scripting.Dictionary")
    Set g_datasDict = CreateObject("Scripting.Dictionary")

    ' PART 3 遍历DB worksheet，并建立字典映射关系
    Const idColx  As Integer = 1 ' 第一列：DataID（固定），注：对应当前DB sheet的列号
    Const groupColx As Integer = 5

    Dim id As String, group As String, prec As String
    Dim inx As Integer

    For inx = 2 To DBWorkSheet.UsedRange.Rows.Count ' i表示行号，因为第一行（DataID、中文、英文） 不需要加入字典，故从2开始
        id = DBWorkSheet.Cells(inx, idColx)
        If Not (id = "") Then
            group = DBWorkSheet.Cells(inx, groupColx)

            ' 创建分组字典
            if group <> "" Then
                g_groupDict(id) = group
            End if

            ' 创建填充数据字典
            If Not g_datasDict.exists(group) Then
                Dim a2D(0 To 1001, 0 To 8) As Variant
                g_datasDict(group) = a2D
            End If
        End If
    Next
End Function


