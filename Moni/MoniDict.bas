Option Explicit
Public g_meanDict, g_precDict, g_visbDict As Object ' ��������ֵ�
' * ��ȡ������Է����ֵ�  k=DataID��v=translated content
' * �� ���д�1��ʼ����
' @ Parameter  DB - string - Ŀ��DB sheet������
Function GetTransDict(DB As String)
    ' PART 1 ����DB worksheet����
    Dim DBWorkSheet As Worksheet
    Set DBWorkSheet = Worksheets(DB)

    ' PART 2 �����ֵ�
    Set g_meanDict = CreateObject("Scripting.Dictionary")
    Set g_precDict = CreateObject("Scripting.Dictionary")
    Set g_visbDict = CreateObject("Scripting.Dictionary")

    ' PART 3 ����DB worksheet���������ֵ�ӳ���ϵ
    Const tagColx  As Integer = 1
    Const meanColx  As Integer = 2
    Const precColx  As Integer = 3
    Const visbColx As Integer = 4

    Dim key As String, mean As String, prec As String, visb As String
    Dim rowx As Integer

    For rowx = 2 To DBWorkSheet.UsedRange.Rows.Count ' i��ʾ�кţ���һ�в���Ҫ�����ֵ䣬�ʴ�2��ʼ
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


