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
Function ParseCsvAndFillCell(resCsv As Variant)
    ' START 禁用自动刷新
    Application.ScreenUpdating = False

    ' PART 1 清空当前sheet的所有单元格内容后
    call ClearCurSheet

    ' PART 2 设置数据对象
    Const precParseColx As Integer = 2 ' 第二列：精度（固定），注：对应当前Parsed sheet的列号
    Const cnParseColx As Integer = 3 ' 第三列：chinese翻译（固定）
    Const enParseColx As Integer = 4 ' 第四列：english（固定）

    Dim idColx As Integer
    idColx = 0

    call CreateSheets(g_groupDict)

    ' PART 3 逐行读取csv文件
    Dim sCurLine As String
    Dim aCsvRowData As Variant
    Dim nCsvCurRowx as Integer
    Open resCsv For Input As #1 ' 打开csv文件，file number #1
    nCsvCurRowx = 0
    Do While Not EOF(1) ' 逐行循环 #1文件
        Line Input #1, sCurLine  ' sCurLine = csv当前行内容（MoldName,SaveDate,Materials,Colour,MoldNum），注：字符串格式
        aCsvRowData = Split(sCurLine, ",") ' 按照逗号，将当前行内容通过Split变为数组

        if nCsvCurRowx > 3 Then
            Dim id as String
            id = aCsvRowData(0)
        End If

        nCsvCurRowx = nCsvCurRowx + 1
    Loop
    Close #1

    ' END
    Application.ScreenUpdating = True  ' 还原屏幕刷新
    ' MsgBox "解析成功！" ' 弹出成功提示

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

' 删除分组sheet
Sub DelGroupSheets()
    Application.DisplayAlerts = False
    Dim nInx as Integer
    ' sheet从1开始计数
    for nInx = 1 To Sheets.Count
        if nInx > 2 Then
            Worksheets(Sheets(3)).Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

' sheet保护（注：无法插入删除行） + 锁定模具头
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
