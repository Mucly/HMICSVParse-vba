Option Explicit

Private Sub ImportCsv_Click()
    ' PART 1
    Call GetTransDict("∂‘’’±Ì")

    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("*.csv, *.csv")
    If csvRes = Empty Then
        Exit Sub
    Else
        Call ParseCsvAndFillCell(csvRes)
    End If

End Sub
