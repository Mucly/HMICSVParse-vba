Option Explicit

Private Sub CommandButton1_Click()
    Dim i As Integer, shp As Variant
    For Each shp In ActiveSheet.Shapes
        If shp.Type = 3 Then
            Debug.Print shp.name
            ActiveSheet.Shapes(shp.name).IncrementLeft 300
        End If
    Next
End Sub

Private Sub ImportCsv_Click()
    ' PART 1
    Call GetTransDict("∂‘’’±Ì")

    Dim csvRes As Variant
    csvRes = Application.GetOpenFilename("*.csv, *.csv")
    If csvRes = Empty Then
        Rem MsgBox "No csv file has been selected!"
        Exit Sub
    Else
        Call ParseCsvAndFillCell(csvRes)
    End If

End Sub

