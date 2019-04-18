' @ 判断文件是否存在的两种方法
' Way One
Dim fso
Set fso = CreateObject("Scripting.Filesystemobject")
If fso.FileExists(filePathName) Then
    '文件存在
Else
    '文件不存在
End if

' Way Two
If dir(filepathname) <>"" Then
    '文件存在
Else
    '文件不存在
End if