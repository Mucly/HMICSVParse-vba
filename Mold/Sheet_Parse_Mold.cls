Option Explicit
' --- click  ---
Private Sub CommandButton1_Click()
    ' PART 1 get Dict
    Dim DBSheet As Object

    Call GetDict(1)

    ' PART 2 file explorer
    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("选择生成的csv文件（*.csv）, *.csv")

    If csvRes = Empty Then
        Exit Sub
    Else
        Call ParseCsvAndFillCell(csvRes)

    End If

    SaveSetting "path", "csvPath", "resCsvPath", csvRes '将源csv文件的完整路径存放到注册表中
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
    Call GetDict(1)

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

    exportFile.writeline (",,,,")

    ' PART 6 Clean Merge-Sheet
    Dim mergeSht As Worksheet: Set mergeSht = Sheets("Merge")
    mergeSht.Visible = True
    Call InitTargetSheet(mergeSht)

    ' PART 7 WriteTitle
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
        Dim sHex As String, vHex As Variant
        sHex = Replace(mergeSht.Cells(rowx, 1), "0x", "&H")
        vHex = CLng(sHex)
        mergeSht.Cells(rowx, serialColx) = vHex
    Next

    ' PART 11 Sort Merge-Sheet Accoring to Serial Colx
    Dim sSerialColx As String: sSerialColx = g_colxAlphaDict(serialColx)
    mergeSht.Activate
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
    Dim sRange As String: sRange = serialAlpha & ":" & serialAlpha
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

    Call GetResCdbPath(1)
    Dim fso
    Set fso = CreateObject("Scripting.Filesystemobject")

    If fso.FileExists(resCdbPathName) Then '判断源cdb文件是否存在
        Dim resCdbPath As String: resCdbPath = resCdbPathName
        Dim bData As Byte

        ' resource
        Open resCdbPath For Binary As #66
        Dim colxx As Integer, row As Variant: row = 5

        ' destination
        Dim dstCdbPath As String, resDir As String, resCsv As String, resCdb As String
        Dim newCdb As String, resName As String
        resDir = GetSetting("csvTools", "MoldToolPath", "resDir")
        resCdb = GetSetting("csvTools", "MoldToolPath", "resCdb")
        resName = GetSetting("csvTools", "MoldToolPath", "resName")
        newCdb = Replace(resCdb, resName, fileName)
        dstCdbPath = resDir + newCdb
        Debug.Print dstCdbPath
        If fso.FileExists(dstCdbPath) Then Kill dstCdbPath ' delete exists dst csb file
        Open dstCdbPath For Binary As #99

        Const maxMoldLen As Integer = 40
        Dim moldName As String: moldName = Range("A4")
        Dim moldNameLen As Integer: moldNameLen = Len(moldName)
        Dim aMoldName(maxMoldLen) As Variant
        Dim item As String, inx As Integer, inxAdd1 As Integer

        For inx = 0 To (maxMoldLen - 1)
            inxAdd1 = inx + 1
            If (inxAdd1 <= moldNameLen) Then aMoldName(inx) = Asc(Mid(moldName, inxAdd1, 1))
        Next

        Dim bytesCnt As Integer: bytesCnt = LOF(66) - 1
        Dim byteX As Integer
        For byteX = 0 To bytesCnt
            Get #66, , bData
            If byteX < 40 Then
                bData = aMoldName(byteX)
            End If

            Put #99, , bData
        Next
        Close #66
        Close #99

    End If
    MsgBox "汇出成功!"

End Sub

Sub GetResCdbPath(inx As Integer)
    Dim resCsvPathName As String, aResCsvPath As Variant, pathNameCnt As Integer, cdbName As String
    Dim sCdbPath As String, pathNum As Integer
    Dim aCdbName As Variant
    resCsvPathName = GetSetting("path", "csvPath", "resCsvPath") '获取注册表中的当前完整路径名

    aResCsvPath = Split(resCsvPathName, "\") '路径按照\分割，将分割后的内容保存到数组中
    pathNameCnt = UBound(aResCsvPath)        '获取当前数组的最大下标

    If resCsvPathName <> "" Then
        cdbName = Replace(aResCsvPath(pathNameCnt), "csv", "cdb") '将csv后缀替换成cdb，得到cdb文件完整名称
        aCdbName = Split(cdbName, ".") '获取源名,如 666.cdb,截取666
        For pathNum = 0 To pathNameCnt - 1    '循环，得到cdb文件完整路径
            resCdbPathName = resCdbPathName + aResCsvPath(pathNum) + "\"
        Next
        SaveSetting "csvTools", "MoldToolPath", "resDir", resCdbPathName    '例：C:\Users\xuwb\Desktop\moldset-cly\
        SaveSetting "csvTools", "MoldToolPath", "resCsv", aResCsvPath(pathNameCnt)   '例：145.csv
        SaveSetting "csvTools", "MoldToolPath", "resCdb", cdbName                    '例：145.cdb
        SaveSetting "csvTools", "MoldToolPath", "resName", aCdbName(0)               '例：145
        resCdbPathName = resCdbPathName + cdbName               '例：C:\Users\xuwb\Desktop\moldset-cly\145.cdb
    End If
End Sub
