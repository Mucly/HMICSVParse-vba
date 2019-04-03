Option Explicit
' TODO 解析时不要固定列号，可以根据关键字来确定列号
' 例如 读到DataID关键字后，获取DataID列号，DataIDValue同理
' --- click 事件 ---
Private Sub CommandButton1_Click()
    ' PART 1 获取多国语言字典
    Dim DBSheet As Object

    Call moldDict.GetDict("对照表")

    ' PART 2 弹出csv文件explorer
    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("选择生成的csv文件（*.csv）, *.csv")
    ' 如果选择了csv文件，则进入解析、翻译、填充工作
    If csvRes = Empty Then
        Exit Sub
    Else
        Call ParseCsvAndFillCell(csvRes)
    End If
End Sub

' 汇出csv文件
Private Sub CommandButton2_Click()
    Call moldDict.GetDict("moldDB")

    ' 获取当前sheet名
    Dim curSheetName As String
    curSheetName = ActiveSheet.Name

    '选择文件下载路径，设置文件名，文件类型
    Dim csvSavepath As Variant
    Dim fileName As String
    Rem fileName = Format(Now(), "_yymmddhhmmss")  ' 防止文件覆盖，故带上当前时间戳
    fileName = Range("A4").Value ' A4单元格固定为模具名
    csvSavepath = Application.GetSaveAsFilename(InitialFileName:=fileName, FileFilter:="(*.csv),*.csv")

    ' fileExplorer按下取消的话，就可以提前结束
    If VarType(csvSavepath) = vbBoolean Then
        Exit Sub
    End If

    'FSO对象引用的后期绑定
    Dim Fs As Object
    Set Fs = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Err_FileOpen
        '创建一个文本文件
        Dim exportFile As Object
        Set exportFile = Fs.createtextfile(csvSavepath)

        Dim idRowx As Integer
        idRowx = 0 ' DataID所在行

        Dim parseRowx As Integer, parseColx As Integer, emptyRows As Integer, maxRows As Integer, maxCols As Integer
        emptyRows = 2 ' 头两行是放按钮的，属于多余行，读取现sheet的单元格时需要多两行
        maxRows = ActiveSheet.UsedRange.Rows.Count + emptyRows
        maxCols = ActiveSheet.UsedRange.Columns.Count

        ' 写入文件内容
        Dim line As String, cellValue As String
        ' 解析后，当前sheet的前面2行是空行，故从第三行开始
        For parseRowx = 3 To maxRows
            line = ""
            For parseColx = 1 To maxCols
                cellValue = Cells(parseRowx, parseColx).Text ' Text 获取单元格实际显示的值
                ' 第一列，则直接取单元格的值
                If parseColx = 1 Then
                    line = cellValue
                ' 其他列，则直接带上先前的数据，并以逗号作为csv文件的分隔符
                Else
                    line = line & "," & cellValue
                End If
            Next
            ' 以csv格式，按行写入数据
            exportFile.writeline (line)
        Next

        ' 关闭所有文件和filesystem object
        exportFile.Close
        Set csvSavepath = Nothing
        Set Fs = Nothing

        MsgBox "Export Success!"
    Exit Sub
    ' 文件已打开的错误提示
Err_FileOpen:
    Dim msg As String, style As String, title As String
    msg = "操作无效！" & Chr(10) & "请关闭如下文件后，再使用此功能" & Chr(10) & csvSavepath & Chr(10)
    style = 16 ' vbCritical
    title = "#Muc " & Err.Description

    Call MsgBox(msg, style, title)
End Sub


