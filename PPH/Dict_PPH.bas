Option Explicit

Function GetTransDict(DB As String)
    ' PART 1 Set Variation
    Dim DBSheet As Worksheet
    Set DBSheet = Worksheets(DB)

    Set g_meanDict = CreateObject("Scripting.Dictionary")

    Const tagColx As Integer = 1
    Const meanColx As Integer = 2

    Dim tag As String, mean As String, rowx As Integer

    ' PART 2 Set Dict's key & value
    For rowx = 2 To DBSheet.UsedRange.rows.Count
        tag = DBSheet.Cells(rowx, tagColx).Value
        If tag <> "" Then
            mean = DBSheet.Cells(rowx, meanColx)
            if mean = "" Then
                g_meanDict(tag) = tag
            Else
                g_meanDict(tag) = mean
            End if
        End If
    Next

End Function
