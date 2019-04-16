Option Explicit
' --- click  ---
Private Sub CommandButton1_Click()
    ' PART 1 get Dict
    Dim DBSheet As Object

    Call GetDict("对照表")

    ' PART 2 file explorer
    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("选择生成的csv文件（*.csv）, *.csv")

    If csvRes = Empty Then
        Exit Sub
    Else
        Call ParseCsvAndFillCell(csvRes)
    End If
End Sub

Sub writeCsvLine(rows As Integer, cols As Integer, offset As Integer, dataSheet As Worksheet, expFile As Object)
    Dim rowx As Integer, colx As Integer, cellValue As String
    Dim sCurLine As String: sCurLine = ""

    For rowx = offset To rows
        For colx = 1 To cols
            cellValue = dataSheet.Cells(rowx, colx).Text
            If colx = 1 Then
                sCurLine = cellValue
            Else
                sCurLine = sCurLine & "," & cellValue
            End If
        Next
        expFile.writeline (sCurLine)
    Next
End Sub

' export file
Private Sub CommandButton2_Click()
    Application.ScreenUpdating = False

    ' PART 2 Get Dict
    Call GetDict("对照表")

    ' PART 3 Pop File-Explorer
    Dim moldHeadSheet As Worksheet: Set moldHeadSheet = Sheets(2)
    Dim fileName As String: fileName = moldHeadSheet.Range("A4")
    Dim csvSavePath As Variant: csvSavePath = Application.GetSaveAsFilename(InitialFileName:=fileName, FileFilter:="(*.csv),*.csv")

    If VarType(csvSavePath) = vbBoolean Then
        Exit Sub
    End If



    ' PART 4 Start Dealing Sheet's Data
    Dim Fs As Object
    Set Fs = CreateObject("Scripting.FileSystemObject")

    Dim exportFile As Object
    Set exportFile = Fs.createtextfile(csvSavePath)
    Dim idRowx As Integer: idRowx = 0
    Dim parseRowx As Integer, parseColx As Integer, emptyRows As Integer, maxRows As Integer, maxCols As Integer
    maxRows = 0 ' vaild cells rows, there has 2 empty rows, so need add 2
    maxCols = 5 ' now, 5 colx is enough

    Dim offsetRows As Integer, colx As Integer

    ' PART 5 Write Mold Head Datas
    maxRows = Application.CountA(moldHeadSheet.Range("A:A")) + 2
    offsetRows = 3
    Call writeCsvLine(maxRows, maxCols, offsetRows, moldHeadSheet, exportFile)

    ' PART 6 Clean Merge-Sheet
    Dim mergeSht As Worksheet: Set mergeSht = Sheets("Merge")
    mergeSht.Visible = True
    Call InitTargetSheet(mergeSht)

    ' PART 7 Write DataID Title To mergeSht
    Dim aTitle As Variant: aTitle = Array("DataID", "DataValue", "Description#1", "Description#2")
    mergeSht.Range("A1").Resize(1, UBound(aTitle) + 1) = aTitle

    ' PART 8 Write Each Group-Sheet's Cells To The Merge-Sheet(hidden)
    ' --- Merge Scattered Sheets To The Merge-Sheet
    Dim inx1 As Integer, MergeOffsetRowx As Variant, groupShtCnts As Integer: groupShtCnts = Sheets.Count
    Dim curSht As Object
    For inx1 = 3 To groupShtCnts
        Set curSht = Sheets(inx1)
        If curSht.Name <> "Merge" Then
            MergeOffsetRowx = Application.CountA(mergeSht.Range("A:A")) + 1
            curSht.UsedRange.offset(1, 0).Copy mergeSht.Cells(MergeOffsetRowx, 1)
        End If
    Next

    ' PART 9 Beauty Merge-Sheet
    With mergeSht.Cells
            .Columns.AutoFit
            .HorizontalAlignment = xlHAlignCenter
    End With

    ' PART 10 Append Serial Colx
    maxRows = Application.CountA(mergeSht.Range("A:A"))
    Dim serialColx As Integer: serialColx = Application.CountA(mergeSht.Range("1:1")) + 1
    Dim rowx As Integer
    For rowx = 2 To maxRows
        Dim sHex As String, nHex As Integer
        sHex = Replace(mergeSht.Cells(rowx, 1), "0x", "&H")
        nHex = CInt(sHex)
        mergeSht.Cells(rowx, serialColx) = nHex
    Next

    ' PART 11 Sort Merge-Sheet Accoring to Serial Colx
    Dim sSerialColx As String: sSerialColx = g_colxAlphaDict(serialColx)
    mergeSht.Columns(sSerialColx & ":" & sSerialColx).Select
    mergeSht.Sort.SortFields.Clear
    mergeSht.Sort.SortFields.Add Key:=Range(sSerialColx & "1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With mergeSht.Sort
        .SetRange Range("A2:" & sSerialColx & maxRows)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' PART 12 Delete Serial Colx
    Dim serialAlpha As String: serialAlpha = g_colxAlphaDict(serialColx)
    Dim sRange As String : sRange = serialAlpha & ":" & serialAlpha
    mergeSht.Columns(sRange).Select
    Selection.Delete Shift:=xlToLeft

    ' PART 13 Continue Writing The CSV Accoring to this Merge Sheet
    offsetRows = 1
    maxRows = Application.CountA(mergeSht.Range("A:A"))
    Call writeCsvLine(maxRows, maxCols, offsetRows, mergeSht, exportFile)

    ' END
    exportFile.Close
    mergeSht.Visible = False

    Set csvSavePath = Nothing
    Set Fs = Nothing

    ' --- Active Parse-Sheet
    Sheets(2).Activate
    Application.ScreenUpdating = True  ' Restore

    ' call ModifyCDB(fileName)

    MsgBox "Export Success!"

End Sub

Sub ModifyCDB(moldName As String)
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & ""

End Sub
