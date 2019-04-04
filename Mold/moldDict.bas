Option Explicit
Public g_id2GroupDict As Object ' { dataID:'mold' }
Public g_dataDict As Object '  { 'groupName' : a2D }
Public g_groupDict as Object ' { 'groupName' : groupName }

' * GetDict  k=DataID，v=translated content
Function GetDict(DB As String)
    ' PART 1 init DBSheet
    Dim DBSheet As Worksheet
    Set DBSheet = Worksheets(DB)

    ' PART 2 Init Dict
    Set g_id2GroupDict = CreateObject("Scripting.Dictionary")
    Set g_dataDict = CreateObject("Scripting.Dictionary")
    Set g_groupDict = CreateObject("Scripting.Dictionary")

    ' PART 3 Travel DB worksheet And Set Dict
    Const idColx  As Integer = 1
    Const groupColx As Integer = 5

    Dim DataID As String, group As String, prec As String,inx As Integer, nDBRows as Integer
    nDBRows = DBSheet.UsedRange.Rows.Count

    For inx = 2 To nDBRows
        DataID = DBSheet.Cells(inx, idColx)
        If DataID <> "" Then
            group = DBSheet.Cells(inx, groupColx)

            if group <> "" Then
                ' create g_id2GroupDict =  { k=DataID, v=group , ...  }
                g_id2GroupDict(DataID) = group

                ' create g_groupDict = { k=group, v=group , ... }
                if Not g_groupDict.exists(group) Then
                    g_groupDict(group) = group
                End If

                ' create dataDict { k=group, v=a2D }
                If Not g_dataDict.exists(group) Then
                    Dim a2D(0 To 1001, 0 To 8) As Variant
                    g_dataDict(group) = a2D
                End If
            End if
        End If
    Next

    ' MoldHead group
    Dim a2DHead(3, 8) As String
    g_dataDict("按钮+模具头") = a2DHead
End Function
