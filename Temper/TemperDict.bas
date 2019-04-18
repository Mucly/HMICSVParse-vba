Option Explicit

Function GetTransDict(DB As String)
    ' PART 1 Set Variation
    Dim DBSheet As Worksheet
    Set DBSheet = Worksheets(DB)

    Set g_meanDict = CreateObject("Scripting.Dictionary")
    Set g_tagPrecDict = CreateObject("Scripting.Dictionary")
    Set g_colxAlphaDict = CreateObject("Scripting.Dictionary")
    Set g_colxPrecDict = CreateObject("Scripting.Dictionary")
    Set g_TempSheetsDict = CreateObject("Scripting.Dictionary")

    Const tagColx As Integer = 1
    Const meanColx As Integer = 2
    Const precColx As Integer = 3

    Dim tag As String
    Dim mean As String, prec As Integer: prec = 0
    Dim rowx As Integer

    ' PART 2 Set Dict's key & value
    For rowx = 2 To DBSheet.UsedRange.rows.Count
        tag = DBSheet.Cells(rowx, tagColx).Value
        If tag <> "" Then
            mean = DBSheet.Cells(rowx, meanColx)
            prec = DBSheet.Cells(rowx, precColx)

            ' tags's meaning, def = tag
            g_meanDict(tag) = mean

            ' tag's prec, if prec not exists, pass
            If rowx > 3 Then
                g_tagPrecDict(tag) = prec
            End If

        End If
    Next

    ' { 1 : A, 2 : B, 3 : C, ... , 27 : AA, 28 : AB, ... }
    Dim myChar As String
    Dim inx As Integer

    For inx = 1 To 208
        Select Case inx
        Case Is <= 26
            g_colxAlphaDict(inx) = Chr(64 + inx)
        Case Is <= 52
            g_colxAlphaDict(inx) = "A" & Chr(64 + inx - 26)
        Case Is <= 78
            g_colxAlphaDict(inx) = "B" & Chr(64 + inx - 52)
        Case Is <= 104
            g_colxAlphaDict(inx) = "C" & Chr(64 + inx - 78)
        Case Is <= 130
            g_colxAlphaDict(inx) = "D" & Chr(64 + inx - 104)
        Case Is <= 156
            g_colxAlphaDict(inx) = "E" & Chr(64 + inx - 130)
        Case Is <= 182
            g_colxAlphaDict(inx) = "F" & Chr(64 + inx - 156)
        Case Else
            g_colxAlphaDict(inx) = "G" & Chr(64 + inx - 182)
        End Select
    Next

End Function

