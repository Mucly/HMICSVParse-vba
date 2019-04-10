For Each st In Worksheets
    If st.Name <> ActiveSheet.Name Then
        st.UsedRange.Offset(1, 0).Copy [a65536].End(xlUp).Offset(1, 0)
    Next
End Sub


Sub 合并当前工作簿下的所有工作表()
'将同一个工作簿中不同工作表合并在一个工作表中（保留原格式不变，非无格式粘贴）。
Application.ScreenUpdating = False
For j = 1 To Sheets.Count
   If Sheets(j).Name <> ActiveSheet.Name Then
       x = ActiveSheet.UsedRange.Rows.Count + 1
       ' Sheets(j).UsedRange.Copy ActiveSheet.Cells(x, 1)
       Sheets(j).UsedRange.offset(1, 0).Copy ActiveSheet.Cells(x, 1)
   End If
Next
Application.ScreenUpdating = True
MsgBox "当前工作簿下的全部工作表已经合并完毕！", vbInformation, "提示"
End Sub

