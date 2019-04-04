Option Explicit
' --- click  ---
Private Sub CommandButton1_Click()
    ' PART 1 get Dict
    Dim DBSheet As Object

    Call moldDict.GetDict("对照表")

    ' PART 2 file explorer
    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("选择生成的csv文件（*.csv）, *.csv")

    If csvRes = Empty Then
        Exit Sub
    Else
        Call ParseCsvAndFillCell(csvRes)
    End If
End Sub

Sub writeCsvLine(rows as Integer, cols as Integer, offset as Integer, dataSheet as Worksheet, expFile as Object)
    Dim rowx as Integer, colx as Integer, cellValue as String
    Dim sCurLine as String : sCurLine = ""

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
    Call moldDict.GetDict("对照表")

    Dim moldHeadSheet As Worksheet : set moldHeadSheet = Sheets(2)

    Dim fileName As String : fileName = moldHeadSheet.Range("A4")
    Dim csvSavepath As Variant : csvSavepath = Application.GetSaveAsFilename(InitialFileName:=fileName, FileFilter:="(*.csv),*.csv")

    If VarType(csvSavepath) = vbBoolean Then
        Exit Sub
    End If

    Dim Fs As Object
    Set Fs = CreateObject("Scripting.FileSystemObject")

    On Error GoTo Err_FileOpen
        Dim exportFile As Object
        Set exportFile = Fs.createtextfile(csvSavepath)
        Dim idRowx As Integer : idRowx = 0
        Dim parseRowx As Integer, parseColx As Integer, emptyRows As Integer, maxRows As Integer, maxCols As Integer
        maxRows = 0 ' vaild cells rows, there has 2 empty rows, so need add 2
        maxCols = 5 ' now, 5 colx is enough

        Dim offsetRows as Integer, colx as Integer

        ' PART 1 write mold head datas
        maxRows = Application.CountA(moldHeadSheet.Range("A:A")) + 2
        offsetRows = 3
        call writeCsvLine(maxRows, maxCols, offsetRows, moldHeadSheet, exportFile)

        ' PART 2 write DataID rows Target
        exportFile.writeline ("DataID,DataValue,,,")

        ' PART 3 write each group sheets datas
        Dim shtKey as Variant
        offsetRows = 2
        For Each shtKey in g_sheetDict
            Dim expSheet as Worksheet : set expSheet = Sheets(shtKey)
            maxRows = Application.CountA(expSheet.Range("A:A"))
            call writeCsvLine(maxRows, maxCols, offsetRows, expSheet, exportFile)
        Next

        exportFile.Close
        Set csvSavepath = Nothing
        Set Fs = Nothing

        MsgBox "Export Success!"
    Exit Sub

Err_FileOpen:
    Dim msg As String, style As String, title As String
    msg = "Failed！" & Chr(10) & "please close those file:" & Chr(10) & csvSavepath & Chr(10)
    style = 16 ' vbCritical
    title = "#Muc " & Err.Description

    Call MsgBox(msg, style, title)
End Sub
