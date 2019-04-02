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
Function ParseCsvAndFillCell(resCsv As Variant, offsetRowx As Integer)
    ' �����Զ�ˢ�£�����ģʽΪ�ֶ�
    Application.ScreenUpdating = False

    ' PART 1 ��յ�ǰsheet�����е�Ԫ�����ݺ󣬹ر���Ļ���£����������ٶ�
    ClearCurSheet

    ' PART 2 �������ݶ���
    Dim fillRowx As Integer
    fillRowx = offsetRowx ' fill the cell start from the third row

    ' PART 3 ���ж�ȡcsv�ļ�
    Dim curLine As String, fillContent As String ' ��ǰ�����ݣ��ַ�����ʽ������������
    ' Range("A3").Resize(rows, colx).Value = a2D

    ' PART 2 ��һ�α���������һ����ά���ݣ�Ϊ�˼ӿ쵥Ԫ������ٶ�
    Dim n2DRows As Integer ' �ļ�������
    Dim n2DCols As Integer
    Open resCsv For Input As #1 ' ��csv�ļ���file number #1
    Do While Not EOF(1) ' ����ѭ�� #1�ļ�
        Line Input #1, curLine

        If n2DCols = 0 Then
            Dim aRowData As Variant
            aRowData = Split(curLine, ",")
            n2DCols = UBound(aRowData)
        End If

        n2DRows = n2DRows + 1

    Loop
    Close #1

    ' PART 3 �ڶ��α���������ά���鸳ֵ
    Dim csvCurRowx As Integer
    Dim a2D(0 To 5001, 0 To 50) As Variant ' ����ֻ�ܳ���
    csvCurRowx = 0
    Open resCsv For Input As #1
    Do While Not EOF(1) ' ����ѭ�� #1�ļ�
        Line Input #1, curLine

        Dim inx As Integer
        For inx = 0 To n2DCols
            Dim aRowData2 As Variant
            Dim nTmp as Integer
            Dim key as String
            ' nTmp = inx + 1
            nTmp = inx
            key = nTmp & ""

            if csvCurRowx = 0 Then
                aRowData2 = Split(curLine, ",")
                if g_meanDict.exists(key) Then
                    aRowData2(inx) = g_meanDict(key)
                End if
                a2D(csvCurRowx, inx) = aRowData2(inx)
            Else ' �������
                curLine = csvCurRowx & curLine
                aRowData2 = Split(curLine, ",")

                Dim prec as Integer
                if g_precDict.exists(key) Then
                    prec = Val(g_precDict(key))
                Else
                    prec = 0
                End if

                If prec = 0 Then
                    a2D(csvCurRowx, inx) = aRowData2(inx)
                Else
                    Dim maxBitWight as Integer, digit as Integer
                    maxBitWight = Application.WorksheetFunction.Power(10, prec)
                    digit = Len(aRowData2(inx)) ' Դ���ֵ�λ��

                    Dim head As String, tail As String, fmt As String   ' ����format�ĸ�ʽ
                    fmt = "General"

                    ' #4 ���� > λ��
                    If prec > digit Then
                        head = "0"
                        tail = String(prec, "0")
                        fmt = head + "." + tail
                    ' ���� < λ��
                    ElseIf prec < digit Then
                        head = String(digit - prec, "0")
                        tail = String(prec, "0")
                        head = "0"
                        fmt = head + "." + tail
                    ' ���� = λ��
                    Else
                        head = "0"
                        tail = String(prec, "0")
                        fmt = head + "." + tail
                    End If   ' #4
                    a2D(csvCurRowx, inx) = Format(aRowData2(inx) / maxBitWight, fmt)
                End If
            End if
        Next

        csvCurRowx = csvCurRowx + 1
    Loop
    Close #1

    Range("A3").Resize(n2DRows + 1, n2DCols + 1) = a2D

    ' END
    Application.ScreenUpdating = True  ' ��ԭ��Ļˢ��
    MsgBox "�����ɹ���" ' �����ɹ���ʾ

End Function

