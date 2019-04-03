Option Explicit
Public g_id2GroupDict As Object ' 资料分组信息 { dataID:'模具' }
Public g_datasDict As Object ' 填充数据分组信息 { 'sheet名字（这里是分组）' : 该sheet行列数据组成的单元格 }
Public g_groupDict as Object ' 分组信息
' * 获取字典  k=DataID，v=translated content
' @Parameter  DB - string - 目标DB sheet的名字
Function GetDict(DB As String)
    ' PART 1 创建DB worksheet对象
    Dim DBWorkSheet As Worksheet
    Set DBWorkSheet = Worksheets(DB)

    ' PART 2 创建多国语言字典
    Set g_id2GroupDict = CreateObject("Scripting.Dictionary")
    Set g_datasDict = CreateObject("Scripting.Dictionary")
    Set g_groupDict = CreateObject("Scripting.Dictionary")

    ' PART 3 遍历DB worksheet，并建立字典映射关系
    Const idColx  As Integer = 1 ' 第一列：DataID（固定），注：对应当前DB sheet的列号
    Const groupColx As Integer = 5

    Dim DataID As String, group As String, prec As String,inx As Integer, nDBRows as Integer
    nDBRows = DBWorkSheet.UsedRange.Rows.Count

    For inx = 2 To nDBRows
        DataID = DBWorkSheet.Cells(inx, idColx)
        If DataID <> "" Then
            group = DBWorkSheet.Cells(inx, groupColx)

            if group <> "" Then
                ' 创建分组对应关系字典{ k=DataID, v=group }
                g_id2GroupDict(DataID) = group

                ' 创建分组字典 { k=group, v=group }
                if Not g_groupDict.exists(group) Then
                    g_groupDict(group) = group
                End If

                ' 创建填充数据字典 { k=group, v=a2D }
                If Not g_datasDict.exists(group) Then
                    Dim a2D(0 To 1001, 0 To 8) As Variant
                    g_datasDict(group) = a2D
                End If
            End if
        End If
    Next
End Function


