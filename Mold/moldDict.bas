Option Explicit
Public g_groupDict As Object
Public g_precDict As Object
Public g_cnDict as Object
Public g_enDict as Object
Public g_sheetDict as Object

' * GetDict  k=DataID，v=translated content
Function GetDict(DB As String)
    ' PART 1 init DBSheet
    Dim DBSheet As Worksheet
    Set DBSheet = Worksheets(DB)

    ' PART 2 Init Dict
    Set g_precDict = CreateObject("Scripting.Dictionary")
    Set g_cnDict = CreateObject("Scripting.Dictionary")
    Set g_enDict = CreateObject("Scripting.Dictionary")
    Set g_groupDict = CreateObject("Scripting.Dictionary")
    Set g_sheetDict = CreateObject("Scripting.Dictionary")

    ' PART 3 Travel DB worksheet And Set Dict
    Const idColx  As Integer = 1
    Const precColx As Integer = 2
    Const cnColx As Integer = 3
    Const enColx As Integer = 4
    Const sheetColx As Integer = 5
    Dim DataID As String, group As String, prec As String, cn As String, en As String, DBRowx As Integer, nDBRows as Variant
    nDBRows = Application.CountA(DBSheet.Range("A:A")) + 1

    For DBRowx = 2 To nDBRows
        DataID = DBSheet.Cells(DBRowx, idColx)
        If DataID <> "" Then
            prec = DBSheet.Cells(DBRowx, precColx)
            cn = DBSheet.Cells(DBRowx, cnColx)
            en = DBSheet.Cells(DBRowx, enColx)
            group = DBSheet.Cells(DBRowx, sheetColx)

            if group <> "" Then
                g_precDict(DataID) = prec
                g_cnDict(DataID) = cn
                g_enDict(DataID) = en
                g_groupDict(DataID) = group

                if Not g_sheetDict.exists(group) Then
                    g_sheetDict(group) = group
                End If

            End if
        End If
    Next
End Function
