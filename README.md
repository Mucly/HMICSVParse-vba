# 操作说明
- 带按钮的sheet，按下按钮，选择res.csv文件
- 名字带DB关键字的sheet用作DB

# 代码说明
## vba密码
    muc
## 类
    Worksheets(sheetName)
        .cells(rowx, colx) ' 返回cell的内容
        .UsedRange.Rows.Count ' 返回当前sheet的有效行数
## 文件操作
    file = Application.GetOpenFileName("description(*.csv), *.csv") ' 调用fileExplorer
    Open file For Input As #1
    Do While Not EOF(1) ' 遍历文件号为1的文件
        Line Input #1, curLine ' 读取该文件的每行内容，并赋值给curLine（字符串）
    Loop
## 单元格配置
    Cells(rowx, colx).Select
    With Selection
        .Width = 20 ' 单元格宽度为20
        .FormulaR1C1(newContent) ' 单元格内容填写
    End With
