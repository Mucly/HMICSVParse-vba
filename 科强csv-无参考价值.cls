Private Sub CommandButton1_Click()
'hs: CSV PDP importer
    Dim PDPImport As Variant
    Dim PDPImportTmp As String
    Cells.Select
    Selection.ClearContents
    Selection.NumberFormat = "General"
    Range("A1").Select
    PDPImport = Application.GetOpenFilename("CSV-Dateien,*.csv,Alle Dateien,*.*")
    If PDPImport = Empty Then Exit Sub

    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 2 'text
    stream.Charset = "utf-8"
    stream.LoadFromFile PDPImport
    Text = stream.ReadText
    stream.Close

    PDPImportTmp = PDPImport & "tmp"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(PDPImportTmp, 2, True, True)
    f.Write Text
    f.Close

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & PDPImportTmp, Destination:= _
        Range("A6"))
        .Name = "PDPImport"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        '.TextFilePlatform = 850 'MS-DOS PC-8
        '.TextFilePlatform = -535 'UTF-8
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1)
        .TextFileThousandsSeparator = " "
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

    Kill PDPImportTmp
End Sub
