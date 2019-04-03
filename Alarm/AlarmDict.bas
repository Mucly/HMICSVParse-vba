Option Explicit
Public g_meanDict As Object

Function GetTransDict(DB As String)
    Dim DBSheet As Worksheet
    Set DBSheet = Worksheets(DB)

    Set g_meanDict = CreateObject("Scripting.Dictionary")
    Const keyColx  As Integer = 1
    Const valueColx As Integer = 2

    Dim k As String,v as String, mean As String
    Dim nRowx As Integer, nRowsCnt As Integer
    nRowsCnt = DBSheet.UsedRange.Rows.Count
    For nRowx = 2 To nRowsCnt
        k = DBSheet.Cells(nRowx, keyColx).Value
        If k <> "" Then
            v = DBSheet.Cells(nRowx, valueColx)
            g_meanDict(k) = v
        End If
    Next
End Function
