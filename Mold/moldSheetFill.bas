Option Explicit
' * ÿ������ǰ��������е�Ԫ�� ---
Function ClearCurSheet()
    Cells.Select
    Selection.ClearContents
End Function
' * ��ȡcsv�ļ��������ֵ䷭����������ǰsheet
' @Parameter  resCsv        - Դcsv�ļ�
'                      offsetRowx - �����ƫ��λ��
' @Remind sheet��������д�1��ʼ
Function ParseCsvAndFillCell(resCsv As Variant)
    ' START �����Զ�ˢ��
    Application.ScreenUpdating = False

    ' PART 1 ��յ�ǰsheet�����е�Ԫ�����ݺ�
    call ClearCurSheet

    ' PART 2 �������ݶ���
    Const precParseColx As Integer = 2 ' �ڶ��У����ȣ��̶�����ע����Ӧ��ǰParsed sheet���к�
    Const cnParseColx As Integer = 3 ' �����У�chinese���루�̶���
    Const enParseColx As Integer = 4 ' �����У�english���̶���

    Dim idColx As Integer
    idColx = 0

    call CreateSheets(g_groupDict)

    ' PART 3 ���ж�ȡcsv�ļ�
    Dim sCurLine As String
    Dim aCsvRowData As Variant
    Dim nCsvCurRowx as Integer
    Open resCsv For Input As #1 ' ��csv�ļ���file number #1
    nCsvCurRowx = 0
    Do While Not EOF(1) ' ����ѭ�� #1�ļ�
        Line Input #1, sCurLine  ' sCurLine = csv��ǰ�����ݣ�MoldName,SaveDate,Materials,Colour,MoldNum����ע���ַ�����ʽ
        aCsvRowData = Split(sCurLine, ",") ' ���ն��ţ�����ǰ������ͨ��Split��Ϊ����

        if nCsvCurRowx > 3 Then
            Dim id as String
            id = aCsvRowData(0)
        End If

        nCsvCurRowx = nCsvCurRowx + 1
    Loop
    Close #1

    ' END
    Application.ScreenUpdating = True  ' ��ԭ��Ļˢ��
    ' MsgBox "�����ɹ���" ' �����ɹ���ʾ

End Function


Sub CreateSheets(groupDict As Object)
    Dim HeadSheet As Worksheet
    Set HeadSheet = Sheets(2)

    ' Dim wsh As Worksheet
    ' Set d = CreateObject("Scripting.Dictionary")
    ' For Each wsh In Worksheets
    '    d(wsh.Name) = ""
    ' Next

    ' If d.exists(group) Then
    '     Application.DisplayAlerts = False
    '     Sheets(group).Delete
    '     Application.DisplayAlerts = True
    ' Else
    '     Sheets.Add After:=DBSheet
    '     ActiveSheet.Name = group
    ' End If
    ' Set d = Nothing

    Call DelGroupSheets()
    Dim aKeys as Variant, nInx as Integer
    aKeys = groupDict.keys

    For nInx = 0 To UBound(aKeys)
        Sheets.Add After:=HeadSheet
        ActiveSheet.Name = aKeys(nInx)
    Next

End Sub

' ɾ������sheet
Sub DelGroupSheets()
    Application.DisplayAlerts = False
    Dim nInx as Integer
    ' sheet��1��ʼ����
    for nInx = 1 To Sheets.Count
        if nInx > 2 Then
            Worksheets(Sheets(3)).Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

' sheet������ע���޷�����ɾ���У� + ����ģ��ͷ
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
