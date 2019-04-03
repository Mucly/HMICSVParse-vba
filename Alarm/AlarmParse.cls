Option Explicit

Private Sub ImportCsv_Click()
    ' PART 1
    Call AlarmDict.GetTransDict("∂‘’’±Ì")

    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("*.csv, *.csv")
    If csvRes = Empty Then
        Rem MsgBox "No csv file has been selected!"
        Exit Sub
    Else
        Dim offsetRowx As Integer
        offsetRowx = 3

        Call AlarmFill.ParseCsvAndFillCell(csvRes, offsetRowx)
    End If

End Sub
