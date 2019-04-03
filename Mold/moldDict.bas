Option Explicit
Public g_id2GroupDict As Object ' ���Ϸ�����Ϣ { dataID:'ģ��' }
Public g_datasDict As Object ' ������ݷ�����Ϣ { 'sheet���֣������Ƿ��飩' : ��sheet����������ɵĵ�Ԫ�� }
Public g_groupDict as Object ' ������Ϣ
' * ��ȡ�ֵ�  k=DataID��v=translated content
' @Parameter  DB - string - Ŀ��DB sheet������
Function GetDict(DB As String)
    ' PART 1 ����DB worksheet����
    Dim DBWorkSheet As Worksheet
    Set DBWorkSheet = Worksheets(DB)

    ' PART 2 ������������ֵ�
    Set g_id2GroupDict = CreateObject("Scripting.Dictionary")
    Set g_datasDict = CreateObject("Scripting.Dictionary")
    Set g_groupDict = CreateObject("Scripting.Dictionary")

    ' PART 3 ����DB worksheet���������ֵ�ӳ���ϵ
    Const idColx  As Integer = 1 ' ��һ�У�DataID���̶�����ע����Ӧ��ǰDB sheet���к�
    Const groupColx As Integer = 5

    Dim DataID As String, group As String, prec As String,inx As Integer, nDBRows as Integer
    nDBRows = DBWorkSheet.UsedRange.Rows.Count

    For inx = 2 To nDBRows
        DataID = DBWorkSheet.Cells(inx, idColx)
        If DataID <> "" Then
            group = DBWorkSheet.Cells(inx, groupColx)

            if group <> "" Then
                ' ���������Ӧ��ϵ�ֵ�{ k=DataID, v=group }
                g_id2GroupDict(DataID) = group

                ' ���������ֵ� { k=group, v=group }
                if Not g_groupDict.exists(group) Then
                    g_groupDict(group) = group
                End If

                ' ������������ֵ� { k=group, v=a2D }
                If Not g_datasDict.exists(group) Then
                    Dim a2D(0 To 1001, 0 To 8) As Variant
                    g_datasDict(group) = a2D
                End If
            End if
        End If
    Next
End Function


