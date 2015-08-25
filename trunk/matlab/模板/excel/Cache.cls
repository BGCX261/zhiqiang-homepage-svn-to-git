VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CITICSCache"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private cacheDict As New scripting.Dictionary   ' 缓存数据
Private mNotCached As String                    '
Private mCacheTimes As Long
Private GetCacheCollection As New Collection
Private mCacheFile


' 获取缓存的个数
Public Property Get count()
    count = cacheDict.count
End Property

' 获取缓存查询次数
Public Property Get CacheTimes()
    CacheTimes = mCacheTimes
End Property

' 没有缓存
Public Property Get NotCached()
    NotCached = mNotCached
End Property

' 设置缓存
Public Sub Add(key, v)
    If (VarType(v) <> vbString Or (Not v Like "*[#/]*")) And (Not v = "") And (Not v = mNotCached) Then
        cacheDict.Item(key) = v
    End If
End Sub

' 获取缓存
Public Property Get Item(key As String)
    mCacheTimes = mCacheTimes + 1
    ' GetCacheCollection.Add key
    If cacheDict.Exists(key) Then
        Item = cacheDict.Item(key)
    Else
        Item = mNotCached
    End If
End Property

' 设置缓存
Public Property Set Item(key As String, v)
    Add key, v
End Property


' 清除缓存，可设置条件，比如ClearCache("*日变现能力")删除所有日变现能力的缓存数据
Public Sub Clear(Optional cond As String = "*")
    If cacheDict Is Nothing Then Exit Sub

    If cond = "*" Then
        cacheDict.RemoveAll
        Exit Sub
    End If
    
    Dim key
    For Each key In cacheDict
        If key Like cond Then cacheDict.Remove key
    Next key
End Sub


' 保存缓存到缓存数据簿
Public Sub Save()
    If Len(mCacheFile) = 0 Then Exit Sub

    Dim key, i As Long
    If mCacheFile Like "*.xls*" Then
        ' 如果缓存文件是Excel文件
        Dim wb As Excel.Workbook, ws As Excel.Worksheet
        
        If (cacheDict Is Nothing) Or (cacheDict.count = 0) Then Exit Sub
        
        Dim eb As New Excel.Application
        If tools.cmd.Exist(mCacheFile) = 1 Then
            Set wb = eb.Application.Workbooks.Open(FileName:=mCacheFile, ReadOnly:=False)
        Else
            Set wb = eb.Application.Workbooks.Add
            wb.Save
        End If
        
        Set ws = wb.Worksheets(1)
        
        Dim val()
        
        ReDim val(1 To cacheDict.count, 1 To 2)
        i = 1
        For Each key In cacheDict
            val(i, 1) = key
            val(i, 2) = cacheDict(key)
            i = i + 1
        Next key
        
        ws.Range("A:B").ClearContents
        ws.Range("A1:B" & cacheDict.count).value = val
        
        wb.Close True
        Set ws = Nothing
        Set wb = Nothing
        Set eb = Nothing
    Else
        ' 如果是其它文件，则以文本格式写入
        i = 1
        
        Open mCacheFile For Output As #1
        
        For Each key In cacheDict
            Print #1, key & "," & VarType(cacheDict(key)) & "," & cacheDict(key)
        Next key
        
        Close #1
    End If
    
    Debug.Print CacheFile & " is saved"
End Sub



' 初始化缓存
Private Sub Class_Initialize()
    mNotCached = "Not Cached"
    mCacheFile = ""
    mCacheTimes = 0
    
    Debug.Print "Cache Initialized"
End Sub

' 类的退出
Private Sub Class_Terminate()
    Save
    Debug.Print "Cache Terminiate " & mCacheFile & ": " & cacheDict.count
End Sub

Public Property Let CacheFile(cf)
    Debug.Print cf & " initialize"
    mCacheFile = cf
    
    If cf Like "*.xls*" Then
        ' 如果缓存文件为Excel文件
        Dim eb As New Excel.Application, wb As Excel.Workbook, ws As Excel.Worksheet, lastColumn As Long, val()
            
        If tools.cmd.Exist(cf) = 1 Then
            Set wb = eb.Application.Workbooks.Open(cf)
        Else
            Set wb = eb.Application.Workbooks.Add
            wb.SaveAs cf
        End If
        
        Set ws = wb.Worksheets(1)
        lastColumn = ws.Range("A10000").End(xlUp).Row
        val = ws.Range("A1:B" & lastColumn).value
        
        If lastColumn > 1 Then
            Dim i As Long
            For i = 1 To lastColumn
                cacheDict.Item(val(i, 1)) = val(i, 2)
            Next i
        End If
        
        wb.Close True
        
        Set ws = Nothing
        Set wb = Nothing
        Set eb = Nothing
    ElseIf cf Like "*.txt" Then
        ' 如果为txt文件，则以文本格式保存
        ' 保存格式为：key,数据类型,数据
        If tools.cmd.Exist(cf) = 1 Then
            Open cf For Input As #1
            
            Dim textLine$, oneItem
            Do While Not EOF(1)
                Line Input #1, textLine
                oneItem = Split(textLine, ",")
                ' 对输入数据进行格式转换
                cacheDict.Item(oneItem(0)) = tools.Convert(oneItem(2), CLng(oneItem(1)))
            Loop
            Close #1
            
            Kill cf
        End If
    End If
    
End Property

Public Property Get CacheFile()
    CacheFile = mCacheFile
End Property


'Function debugCache()
'    Dim val(), i
'    ReDim val(1 To GetCacheCollection.Count, 1 To 1)
'
'    With GetCacheCollection
'        For i = 1 To .Count
'            val(i, 1) = .item(i)
'        Next i
'    End With
'
'    ThisWorkbook.Worksheets("缓存").Range("C1:C" & GetCacheCollection.Count) = val
'End Function

