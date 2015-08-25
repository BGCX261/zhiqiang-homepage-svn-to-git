VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CITICSCMD"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private fso As New scripting.FileSystemObject

' 返回一个资源定位串的路径部分
Function Path$(uri$)
    Dim i&
    
    i = InStrRev(uri, "\")
    Path = mid(uri, 1, i)
End Function

' 返回一个资源定位串的文件名
Function FileName$(uri$)
    Dim i&, j&
    
    i = InStrRev(uri, "\")
    i = InStrRev(uri, "/")
    FileName = mid(uri, IIf(i < j, j, i) + 1, 10000)
End Function

' 判断一个资源定位串对应的文件或者文件夹是否存在，
' 返回1表示文件，-1表示文件夹，0表示不存在
Function Exist(ByVal uri$) As Long
    If fso.FileExists(uri) Or (InStr(uri, "*") > 0) Or (InStr(uri, "?") > 0) Then
        Exist = 1
    ElseIf fso.FolderExists(uri) Then
        Exist = -1
    Else
        Exist = 0
    End If
End Function

' 新建文件夹
Function Mkdir(uri$)
    If Exist(uri) = 0 Then
        fso.CreateFolder uri
    End If
End Function

' 文件大小
Function Size(uri$)
    Dim fl As File
    Set fl = fso.GetFile(uri)
    If Exist(uri) = 1 Then
        Size = fl.Size
    End If
End Function

' 获得文件最后修改时间
Function LastModifyDate(uri$)
    Dim fl As File
    Set fl = fso.GetFile(uri)
    If Exist(uri) = 1 Then
        LastModifyDate = fl.DateLastModified
    End If
End Function

' 复制文件或者文件夹
Function Copy(source$, destination$, Optional overwrite As Boolean = True)
    Select Case Exist(source)
        Case 1
            fso.CopyFile source, destination, overwrite
        Case -1
            If Exist(destination) = 0 Then
                Mkdir destination
            End If
            fso.CopyFolder source, destination, overwrite
        Case 0
            Debug.Print " " & source & " is not found"
    End Select
End Function

' 移动文件或文件夹
Function Move(source$, destination$)
    Select Case Exist(source)
        Case 1
            fso.CopyFile source, destination
        Case -1
            fso.CopyFolder source, destination
        Case 0
            Debug.Print " " & source & " is not found"
    End Select
End Function

' 删除文件或文件夹
Function Delete(uri$)
    Select Case Exist(uri)
        Case 1
            fso.DeleteFile uri
        Case -1
            fso.DeleteFolder uri
        Case 0
            Debug.Print " " & uri & " is not found"
    End Select
End Function

Function Open_(uri$)
    Select Case Exist(uri)
        Case 1
            Shell "cmd /c """ & uri & """"
        Case -1
            Shell "explorer " & uri, vbNormalFocus
        Case 0
            Debug.Print " " & uri & " is not found"
    End Select
End Function

' 从文本文件读入所有的文字
Public Function ReadTextFile(ByVal i_FileName As String) As String
    Dim myFNo&  'file number to open file
    Dim myText As String  'string where text gets read into
    
    'only open file if it can be found
    If Dir(i_FileName, vbNormal) <> "" Then
        'obtain the next available file number
        myFNo = FreeFile
        'open file in binary mode
        Open i_FileName For Binary As #myFNo
        'initialize string to receive text with
        'as many spaces as the file has bytes
        myText = Space(LOF(myFNo))
        'read everything at once
        Get #myFNo, , myText
        'close file
        Close #myFNo
        'return text
        ReadTextFile = myText
    End If
End Function
