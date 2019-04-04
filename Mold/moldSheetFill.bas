Option Explicit
' clear current sheet's cells
Function ClearCurSheet()
    Cells.Select
    Selection.ClearContents
End Function

Function ParseCsvAndFillCell(resCsv As Variant)
    ' START
    Application.ScreenUpdating = False

    ' PART 1
    Call ClearCurSheet

    ' PART 2
    call CreateGroupSheets(g_groupDict)

    ' PART 3
    Call SetDataDict(resCsv, g_dataDict)

    ' END
    Application.ScreenUpdating = True  ' Restore
    ' MsgBox "Success！"

End Function
' TODO
Function GetArrayVaildCnt(a2D As Variant)
    Dim rowx As Integer, colx As Integer, nFillRowx As Integer
    nFillRowx = 0
    colx = 0
    For rowx = 0 To UBound(a2D)
        If a2D(rowx, 0) <> "" Then
            nFillRowx = nFillRowx + 1
        End If
    Next
    GetArrayVaildCnt = nFillRowx
End Function

Sub SetDataDict(resCsv As Variant, DataDict As Object)
    ' PART 3 Read Csv By Line, Then Set Each Group's a2D in DataDict
    Dim sCurLine As String
    Dim aCsvRowData As Variant
    Dim nCsvCurRowx As Integer
    Open resCsv For Input As #1 ' csv.fileNumber == #1
    nCsvCurRowx = 1
    Do While Not EOF(1)
        Line Input #1, sCurLine
        aCsvRowData = Split(sCurLine, ",")

        Dim colx As Integer, fillColx as Integer, cellValue As String, DataID As String, group As String
        DataID = aCsvRowData(0) : fillColx = 0 :

        if nCsvCurRowx < 4 Then
            Debug.print ""
        Else
            If g_id2GroupDict.exists(DataID) Then
                Dim fillSheet As Worksheet, fillRowx as Integer
                group = g_id2GroupDict(DataID)
                Set fillSheet = Sheets(group)
                ' fillRowx = fillSheet.Range("A65536").End(xlUp).Row + 1
                fillRowx = Application.CountA(fillSheet.Range("A:A")) + 1

                For colx = 0 To UBound(aCsvRowData)

                    fillColx = colx + 1
                    cellValue = aCsvRowData(colx)
                    ' the top two lines's content is MoldHeader
                    Debug.print group, fillRowx, fillColx, cellValue
                    fillSheet.Cells(fillRowx, fillColx) = cellValue
                Next
            End if
        End if
        nCsvCurRowx = nCsvCurRowx + 1
    Loop
    Close #1
End Sub

Sub CreateGroupSheets(groupDict As Object)
    Dim HeadSheet As Worksheet
    Set HeadSheet = Sheets(2)

    Call DelGroupSheets
    Dim aKeys As Variant, nInx As Integer
    aKeys = groupDict.keys

    For nInx = 0 To UBound(aKeys)
        Sheets.Add After:=HeadSheet
        ActiveSheet.Name = aKeys(nInx)
        Dim aTitle as Variant : aTitle = Array("DataID", "DataValue", "中文翻译", "English")
        ActiveSheet.Range("A1").Resize(1, UBound(aTitle) + 1) = aTitle
    Next
End Sub

Sub DelGroupSheets()
    Application.DisplayAlerts = False
    Dim nInx As Integer
    ' sheet's index start from 1
    For nInx = 1 To Sheets.Count
        If nInx > 2 Then
            ' the top two sheets is standard, delete others sheets only
            Worksheets(Sheets(3)).Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

' sheet Protect
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
