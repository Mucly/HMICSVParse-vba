Option Explicit
' --- click  ---
Private Sub CommandButton1_Click()
    ' PART 1
    Dim DBSheet As Object

    Call MoniDict.GetTransDict("∂‘’’±Ì")

    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("*.csv, *.csv")
    If csvRes = Empty Then
        Rem MsgBox "No csv file has been selected!"
        Exit Sub
    Else
        Dim offsetRowx As Integer
        offsetRowx = 3

        Call sheetFill.ParseCsvAndFillCell(csvRes, offsetRowx)
    End If

End Sub