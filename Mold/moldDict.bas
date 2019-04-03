Option Explicit
Public g_groupDict As Object ' ���Ϸ�����Ϣ { dataID:'ģ��' }
Public g_datasDict As Object ' ������ݷ�����Ϣ { 'sheet���֣������Ƿ��飩' : ��sheet����������ɵĵ�Ԫ�� }
' * ��ȡ������Է����ֵ�  k=DataID��v=translated content
' @Parameter  DB - string - Ŀ��DB sheet������
Function GetTransDict(DB As String)
    ' PART 1 ����DB worksheet����
    Dim DBWorkSheet As Worksheet
    Set DBWorkSheet = Worksheets(DB)

    ' PART 2 ������������ֵ�
    Set g_groupDict = CreateObject("Scripting.Dictionary")
    Set g_datasDict = CreateObject("Scripting.Dictionary")

    ' PART 3 ����DB worksheet���������ֵ�ӳ���ϵ
    Const idColx  As Integer = 1 ' ��һ�У�DataID���̶�����ע����Ӧ��ǰDB sheet���к�
    Const groupColx As Integer = 5

    Dim id As String, group As String, prec As String
    Dim inx As Integer

    For inx = 2 To DBWorkSheet.UsedRange.Rows.Count ' i��ʾ�кţ���Ϊ��һ�У�DataID�����ġ�Ӣ�ģ� ����Ҫ�����ֵ䣬�ʴ�2��ʼ
        id = DBWorkSheet.Cells(inx, idColx)
        If Not (id = "") Then
            group = DBWorkSheet.Cells(inx, groupColx)

            ' ���������ֵ�
            if group <> "" Then
                g_groupDict(id) = group
            End if

            ' ������������ֵ�
            If Not g_datasDict.exists(group) Then
                Dim a2D(0 To 1001, 0 To 8) As Variant
                g_datasDict(group) = a2D
            End If
        End If
    Next
End Function


