VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CITICSDatabase"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private sqlDict As New scripting.Dictionary   ' 缓存数据
Private Sub Class_Initialize()
    With sqlDict
        .Add "聚源", conJYDBInfo
        .Add "jydb", conJYDBInfo
        .Add "yzyq", conGoGoalInfo
        .Add "一致预期", conGoGoalInfo
        .Add "wind", conWindInfo
        .Add "万德", conWindInfo
        .Add "xrisk", conxRiskInfo
        .Add "xrisk33", conxRisk33Info
        .Add "gpjy", conGPJYInfo
        .Add "高频交易", conGPJYInfo
    End With
End Sub


' 查询数据库，返回RecordSet对象
' sql: 数据库查询语句
' sqlInfo: 数据库连接信息，或者直接指定数据库，比如"Wind"、"JYDB"等，利用内设的数据库连接信息
Public Function Query(sql As String, Optional ByVal sqlInfo As String = "xRisk") As ADODB.Recordset
    sqlInfo = GetSQLInfo(sqlInfo)
    Dim cnn As New ADODB.Connection, rst As New ADODB.Recordset
    On Error GoTo errorhander
    If cnn.ConnectionString <> "" Then cnn.Close
    cnn.Open sqlInfo
    rst.Open sql, cnn
    Set Query = rst
errorhander:
End Function

' 查询数据库，返回单个数值
' sql: 数据库查询语句
' sqlInfo: 数据库连接信息，或者直接指定数据库，比如"Wind"、"JYDB"等，利用内设的数据库连接信息
Public Function QueryOne(sql As String, _
                Optional ByVal sqlInfo As String = "xRisk")
    sqlInfo = GetSQLInfo(sqlInfo)
    
    ' Debug.Assert Not sqlInfo Like "*gogoal*"
    Static cnn As New ADODB.Connection, rst As New ADODB.Recordset
    On Error GoTo errorhander
    cnn.Open sqlInfo
    On Error GoTo errorhander
    rst.Open sql, cnn
    On Error GoTo errorhander
    QueryOne = rst.Fields.Item(0).value
    
    ' rst.Close
errorhander:
    If Err.Number > 0 Then QueryOne = "#N/A"
        
    On Error Resume Next
    
    cnn.Close
End Function

' 查询数据库，返回单个数值
' sql: 数据库查询语句
' sqlInfo: 数据库连接信息，或者直接指定数据库，比如"Wind"、"JYDB"等，利用内设的数据库连接信息
Public Function QueryToCell(sql$, Optional rng As Excel.Range, Optional ByVal sqlInfo$ = "xRisk")
    sqlInfo = GetSQLInfo(sqlInfo)
    Static cnn As New ADODB.Connection, rst As New ADODB.Recordset
    ' MsgBox sqlInfo
    cnn.Open sqlInfo
    rst.Open sql, cnn

    rng.CopyFromRecordset rst
    
    rst.Close
    cnn.Close
End Function

' 查询存储过程
Public Function QueryStoredProc(procName$, para, Optional returnPara As Boolean = True) As ADODB.Recordset
    Dim cnn As New ADODB.Connection, com As New ADODB.Command
    On Error GoTo errorhander
    cnn.Open conxRiskInfo
    On Error GoTo errorhander
    
    With com
        .ActiveConnection = cnn
        .CommandType = adCmdStoredProc

        .CommandText = procName

        ' 获取存储过程的参数定义
        .Parameters.Refresh
        
        ' 如果存在输出参数，则删除它，默认第一个为输出参数
        On Error Resume Next
        If returnPara Then .Parameters.Delete 0
        
        ' 设置输入参数的值
        If IsArray(para) Then
            Dim i
            For i = 0 To UBound(para)
                .Parameters.Item(i).value = para(i)
            Next i
        End If


        ' 改变输入参数大小
        Dim tmpp
        For Each tmpp In .Parameters
            tmpp.Size = 255
        Next tmpp

        ' 获取参数返回值
        Set QueryStoredProc = .Execute()
    End With
    
errorhander:
    If Err.Number > 0 Then Set QueryStoredProc = Nothing
    ' cnn.Close
End Function


Private Function GetSQLInfo(source)
    If sqlDict.Exists(LCase(source)) Then
        GetSQLInfo = sqlDict.Item(LCase(source))
    Else
        GetSQLInfo = source
    End If
End Function

