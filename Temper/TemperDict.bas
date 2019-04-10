Option Explicit

Function GetTransDict(DB As String)
    ' PART 1 Set Variation
    Dim DBSheet As Worksheet
    Set DBSheet = Worksheets(DB)

    Set g_meanDict = CreateObject("Scripting.Dictionary")
    Set g_precDict = CreateObject("Scripting.Dictionary")
    Set g_colxAlphaDict = CreateObject("Scripting.Dictionary")

    Const tagColx As Integer = 1
    Const meanColx As Integer = 2
    Const precColx As Integer = 3

    Dim tag As String
    Dim mean As String, prec as Integer : prec = 0
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
                g_precDict(tag) = prec
            End If

        End If
    Next

    ' { 1 : A, 2 : B, 3 : C, ... }
    Dim myChar As String
    Dim inx As Integer
    For inx = 1 To 27
        g_colxAlphaDict(inx) = chr(64 + inx)
    Next

End Function