Option Explicit
' * GetDict  k=DataID，v=translated content
Function GetDict(shtInx As Integer)
    ' PART 1 init DBSheet
    Dim DBSheet As Worksheet
    Set DBSheet = Sheets(shtInx)

    ' PART 2 Init Dict
    Set g_precDict = CreateObject("Scripting.Dictionary")
    Set g_cnDict = CreateObject("Scripting.Dictionary")
    Set g_enDict = CreateObject("Scripting.Dictionary")
    Set g_groupDict = CreateObject("Scripting.Dictionary")
    Set g_sheetDict = CreateObject("Scripting.Dictionary")
    Set g_colxAlphaDict = CreateObject("Scripting.Dictionary")

    ' PART 3 Travel DB worksheet And Set Dict
    Const idColx  As Integer = 1
    Const precColx As Integer = 2
    Const cnColx As Integer = 3
    Const enColx As Integer = 4
    Const sheetColx As Integer = 5
    Dim DataID As String, group As String, prec As String, cn As String, en As String, DBRowx As Integer, nDBRows As Variant
    nDBRows = Application.CountA(DBSheet.Range("A:A")) + 1

    For DBRowx = 2 To nDBRows
        DataID = DBSheet.Cells(DBRowx, idColx)
        If DataID <> "" Then
            prec = DBSheet.Cells(DBRowx, precColx)
            cn = DBSheet.Cells(DBRowx, cnColx)
            en = DBSheet.Cells(DBRowx, enColx)
            group = DBSheet.Cells(DBRowx, sheetColx)

            If group <> "" Then
                g_precDict(DataID) = prec
                g_cnDict(DataID) = cn
                g_enDict(DataID) = en
                g_groupDict(DataID) = group

                If Not g_sheetDict.exists(group) Then
                    g_sheetDict(group) = group
                End If

            End If
        End If
    Next

    ' { 1 : A, 2 : B, 3 : C, ... }
    Dim myChar As String
    Dim inx As Integer
    For inx = 1 To 27
        g_colxAlphaDict(inx) = Chr(64 + inx)
    Next

End Function

