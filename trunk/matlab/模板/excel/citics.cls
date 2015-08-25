VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CITICSTools"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public cmd As New CITICSCMD

' 将目前活动sheet保存为pdf文件，指定文件名地址为FileName
'
' author: zhiqiang@citics, 2009
Public Sub saveAsPDF(ws As Excel.Worksheet, FileName As String)
    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
            FileName, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

' 利用Outlook发送邮件
Public Function sendMailWithOutlook( _
        Optional receiptions As String = "zhangzq@citics.com; mathzqy@gmail.com; zhang@zhiqiang.org", _
        Optional subject As String, _
        Optional body As String, _
        Optional attachments, _
        Optional autoSend As Boolean = False, _
        Optional rng As Excel.Range)
        
    Static xOutLook  As Outlook.Application
    Dim xMail As Outlook.MailItem
    
    Set xOutLook = GetObject(, "Outlook.Application")
    Set xMail = xOutLook.CreateItem(olMailItem)
    With xMail
        .Display
        Dim signature As String
        signature = .HTMLBody
        .To = receiptions
        .subject = subject
        .HTMLBody = body
        
        If Not IsMissing(rng) Then
            ' .HTMLBody = RangetoHTML(rng)
            Dim xDoc
            Set xDoc = xMail.Application.ActiveInspector.WordEditor
            rng.Copy
            xDoc.Range.PasteSpecial DataType:=wdPasteBitmap, Placement:=wdInLine
        End If
        ' .Importance = olImportanceHigh ' 设置优先级
        If IsArray(attachments) Then
            Dim attachment
            For Each attachment In attachments
                .attachments.Add attachment
            Next attachment
        End If
        
        .HTMLBody = .HTMLBody & signature
        
        If autoSend Then
            .Send
        Else
            .Display
        End If
    End With
End Function


' 通过Lotus发送邮件
Public Function sendMailWithLotus(vaRecipient As Variant, emailTitle, _
        emailBody, vaFiles As Variant, Optional sentOut = False, Optional sheetRange = "")
    Dim noSession    As Object
    Dim noDatabase   As Object
    Dim noDocument   As Object
    Dim noAttachment As Object
    Dim richTextBody As Object
    Dim ws As Object
    Dim i            As Long
    Const EMBED_ATTACHMENT = 1454
    
    ' 如果需要手动选取附件，保留下面一行语句
    ' vaFiles = Application.GetOpenFilename(FileFilter:="Excel Filer (*.xls),*.xls", _
              Title:="Attach files for outgoing E_Mail", MultiSelect:=True)
    If Not IsArray(vaFiles) Then Exit Function
    
    'Insert Lotus Notes COM object.
    Set noSession = CreateObject("Notes.NotesSession")
    Set ws = CreateObject("Notes.NotesUIWorkspace")

    Set noDatabase = noSession.GETDATABASE("", "")
    If noDatabase.IsOpen = False Then noDatabase.OPENMAIL
    Set noDocument = noDatabase.createdocument
    Set noAttachment = noDocument.CreateRichTextItem("attachment")
    Set richTextBody = noDocument.CreateRichTextItem("Body")
    If IsArray(vaFiles) Then
        With noAttachment
             For i = LBound(vaFiles) To UBound(vaFiles)
                 .EmbedObject EMBED_ATTACHMENT, "", vaFiles(i)
             Next i
        End With
    End If

    With noDocument
         .Form = "Memo"
         .SendTo = vaRecipient
         .subject = emailTitle
         .body = emailBody
         .SAVEMESSAGEONSEND = True
         .PostedDate = Now() - 100
'         .SEND 0, vaRecipient
    End With
    
    On Error Resume Next
    If Not (sheetRange = "") Then
        Set uidoc = ws.EDITDOCUMENT(True, noDocument)
        
        sheetRange.Copy ' Picture Appearance:=xlScreen, Format:=xlPicture
        
        Call uidoc.GOTOFIELD("Body")
        Call uidoc.Paste  ' CAN I PASTE SPECIAL
'        uidoc.Close False
        Call uidoc.Save

'        Dim tempObject As Object
'        Set noDocument = uidoc.DOCUMENT
'       Call noDocument.Save(True, True)
    End If
    If sentOut Then
        noDocument.Send False, vaRecipient
    End If
    
    Set noDocument = Nothing
    Set noDatabase = Nothing
    Set noSession = Nothing
    Set ws = Nothing
    Set tempObject = Nothing
    Set uidoc = Nothing
    Set richTextBody = Nothing
    ' MsgBox "This file is send  OK", vbInformation
End Function



Public Function match(value, ByVal arr, matchType As Long)
    Dim mid As Long, first As Long, last As Long, i As Long
    
    match = -1
    
    ' 精确查找
    If matchType = 0 Then
        For i = LBound(arr, 1) To UBound(arr, 2)
            If arr(i, 1) = value Then
                match = i + LBound(arr, 1)
            End If
        Next i
        
        If match < 0 Then match = "#Value!"
        
        Exit Function
    
    ' 查找数组中大于等于value的最小值，要求value是排好序的（如果arr第一项是字符，则不参与查找）
    ElseIf matchType = -1 Then
        ' 如果arr第一项是字符，则不参与查找（此项主要是为了方便使用）
        first = IIf(VarType(arr(1, 1)) = vbString, 2, 1)
        
        ' 如果最后一项都小于value，则返回查找错误
        last = UBound(arr, 1)
        If arr(last, 1) < value Then
            match = "#Value!"
            Exit Function
        End If
        
        ' 进行经典的二分查找
        Do While (last > first + 1)
            mid = Int((last + first) / 2)
            If arr(mid, 1) = value Then
                match = mid
                Exit Function
            ElseIf arr(mid, 1) < value Then
                first = mid + 1
            Else
                last = mid
            End If
        Loop
        
        If arr(first) < value Then
            match = first + 1
        Else
            match = first
        End If
    End If
    
End Function

Public Function StopRefresh(Optional isStop, Optional ea As Excel.Application)
    Static screenUpdateState, statusBarState, calcState As Excel.XlCalculation, _
        eventsState, displayPageBreaksState, isStopRefresh
    
    If IsMissing(isStop) Then
        If isStopRefresh Then isStop = 0 Else isStop = 1
    End If
    
    If ea Is Nothing Or IsMissing(ea) Then
        Set ea = GetExcelApp()
    End If
    
    If (isStop) = True Then
        If Not isStopRefresh = True Then
            ' 如果没有进入"节能模式"
            isStopRefresh = True
            
            '获得当前的Excel设置状态,将其放置在代码的开头
            ' 但是当已经进入节能模式时，以下代码可能会冲掉之前的设置，所以以下代码只能在
            ' 之前在非节能模式下时才可使用
            screenUpdateState = ea.ScreenUpdating
            statusBarState = ea.DisplayStatusBar
            ' calcState = ea.Calculation
            
            eventsState = ea.EnableEvents
            ' displayPageBreaksState = ea.ActiveSheet.DisplayPageBreaks '注:这是工作表级的设置
        End If
        
    
        '关闭一些Excel功能使代码运行更快
        ea.ScreenUpdating = False
        ea.DisplayStatusBar = False
        ' ea.Calculation = xlCalculationManual
        ea.EnableEvents = False
        ' ea.ActiveSheet.DisplayPageBreaks = False '注:这是工作表级的设置
    ElseIf (isStop) = False Then
        ' 退出节能模式
        isStopRefresh = False
        
        '代码运行后,恢复Excel原来的状态;将下面的代码放在代码的末尾
        ' ea.ScreenUpdating = screenUpdateState
'        ea.DisplayStatusBar = statusBarState
'        ea.Calculation = calcState
'        ea.EnableEvents = eventsState
'        ea.ActiveSheet.DisplayPageBreaks = displayPageBreaksState '注:这是工作表级的设置
    Else
        isStopRefresh = False
        ea.ScreenUpdating = True
        ea.DisplayStatusBar = True
        ea.Calculation = xlCalculationAutomatic
        ea.EnableEvents = True
        ea.ActiveSheet.DisplayPageBreaks = True
    End If
End Function

' 刷新数据透视表的数据源
Public Function RefreshPivotCache(pt As Excel.PivotTable, Optional latestData = 0, Optional f As String)
    Dim pf As Excel.PivotField
    If Not latestData = 0 Then
        Set pf = pt.PivotFields(f)
        With pf
            If .PivotItems(.PivotItems.count) < latestData Then
                pt.PivotCache.Refresh
            End If
        End With
    Else
        pt.PivotCache.Refresh
    End If
    
    Set pf = Nothing
End Function

' 列数转换为字母
Public Function Num2Alphabat(i As Long) As String
    Num2Alphabat = Chr(i + Asc("A") - 1)
End Function

' 将
Function RangetoHTML(rng As Excel.Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2007
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Excel.Workbook
 
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
 
    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = ea.Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        ea.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
 
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).name, _
         source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
 
    'Close TempWB
    TempWB.Close savechanges:=False
 
    'Delete the htm file we used in this function
    Kill TempFile
 
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function



' 将val转化为对应类型
Function Convert(val, valueType As VbVarType)
    If valueType = vbBoolean Then
        Convert = IIf(VarType(val) = vbBoolean, val, CBool(val))
    ElseIf valueType = vbDate Then
        Convert = CDate(val)
    ElseIf valueType = vbDouble Or valueType = vbSingle Or valueType = vbDecimal Then
        Convert = CDbl(val)
    ElseIf valueType = vbString Then
        Convert = IIf(VarType(val) = vbString, val, CStr(val))
    ElseIf valueType = vbLong Or valueType = vbInteger Then
        Convert = IIf(VarType(val) = vbLong, val, CLng(val))
    Else
        Convert = val
    End If
End Function

' 修改Excel数据
Sub ChangeODBCConnection(wb As Excel.Workbook, conName As String, conText As String, _
        Optional comText As String = "")
    
    With wb.Connections(conName).ODBCConnection
        If Len(comText) Then .CommandText = comText
        If Len(conText) Then .Connection = conText
    End With
    wb.Connections(conName).Refresh
End Sub


' 根据指定区域，将空单元格所在行隐藏
Public Function HideEmptyRow(Optional sel As Excel.Range)
    Dim theCell As Excel.Range
    
    If sel Is Nothing Then
        Dim ea As Excel.Application
        Set ea = GetObject(, "Excel.Application")
        Set sel = ea.Selection
    End If

    For Each theCell In sel
        If theCell.value = "" Then
            If Not theCell.EntireRow.Hidden Then theCell.EntireRow.Hidden = True
        Else
            If theCell.EntireRow.Hidden Then theCell.EntireRow.Hidden = False
        End If
    Next theCell
End Function

' 将字符串按照指定的格式转化为日期
Public Function ToDate(ByVal dt$, ByVal fm$) As Date
    Dim tmp$
    dt = VBA.Trim(dt)
    fm = VBA.LCase(VBA.Trim(fm))
    
    Dim length&, start&, i&, y$, m$, d$
    
    ' 提取年份
    start = 0
    length = 0
    For i = 1 To Len(dt)
        If VBA.mid(fm, i, 1) = "y" Then
            length = length + 1
            If start = 0 Then start = i
        End If
    Next i
    y = VBA.mid(dt, start, length)
    If Len(y) = 2 Then y = "20" & y
    ' 提取月份
    start = 0
    length = 0
    For i = 1 To Len(dt)
        If VBA.mid(fm, i, 1) = "m" Then
            length = length + 1
            If start = 0 Then start = i
        End If
    Next i
    m = VBA.mid(dt, start, length)
    ' 提取日期
    start = 0
    length = 0
    For i = 1 To Len(dt)
        If VBA.mid(fm, i, 1) = "d" Then
            length = length + 1
            If start = 0 Then start = i
        End If
    Next i
    d = VBA.mid(dt, start, length)
    ' 合成字符串，转换成日期格式
    ToDate = CDate(y & "-" & m & "-" & d)
End Function

' 根据列数提取列的字母
Function ColumnLetter(ByVal ColNumber&) As String
    If ColNumber < 0 Or ColNumber > 26 * 26 + 26 Then
        Debug.Print "Error column number: " & ColNumber
    ElseIf ColNumber <= 26 Then
        ColumnLetter = VBA.Chr(ColNumber + VBA.Asc("A") - 1)
    ElseIf ColNumber <= 26 * 26 + 26 Then
        ColNumber = ColNumber - 27
        ColumnLetter = VBA.Chr(Int(ColNumber / 26) + VBA.Asc("A")) & VBA.Chr((ColNumber Mod 26) + VBA.Asc("A"))
    Else
        ColumnLetter = VBA.Chr(Int(ColNumber / 26 / 26) + VBA.Asc("A") - 1) & VBA.Chr((Int(ColNumber / 26) Mod 26) + VBA.Asc("A") - 1) & VBA.Chr((ColNumber Mod 26) + VBA.Asc("A") - 1)
    End If
End Function

' 根据列的字母提取列的列数
Function ColumnNumber(ColLetter$) As Long
    Select Case Len(ColLetter)
        Case 1
            ColumnNumber = VBA.Asc(ColLetter) - VBA.Asc("A") + 1
        Case 2
            ColumnNumber = 26 + (VBA.Asc(VBA.mid(ColLetter, 1, 1)) - VBA.Asc("A")) * 26 + (VBA.Asc(VBA.mid(ColLetter, 2, 1)) - VBA.Asc("A") + 1)
       ' Case 3
       '     ColumnNumber = (VBA.Asc(VBA.mid(ColLetter, 1, 1)) - VBA.Asc("A") + 1) * 26 * 26 + (VBA.Asc(VBA.mid(ColLetter, 2, 1)) - VBA.Asc("A") + 1) * 26 + (VBA.Asc(VBA.mid(ColLetter, 3, 1)) - VBA.Asc("A") + 1)
        Case Else
            Debug.Print "Error column letters: " & ColLetter
    End Select
End Function

' 将公式转化为数值
Function FormulasToValues(rng As Excel.Range)
    rng.value = rng.value
End Function


Function StringToNumber(rng As Excel.Range)
    Dim cl
    For Each cl In rng
        If VarType(cl.value) = vbString Then
            On Error Resume Next
            cl.value = val(cl.value) * 1
            If Err.Number > 0 Then
                Err.Clear
            Else
                Exit Function
            End If
            On Error Resume Next
            cl.value = val(cl.value) * 1
            If Err.Number > 0 Then
                Err.Clear
            Else
                Exit Function
            End If
            On Error Resume Next
            cl.value = val(cl.value) * 1
            If Err.Number > 0 Then
                Err.Clear
            Else
                Exit Function
            End If
        End If
    Next cl
End Function


Public Sub RegFun()
    ' Call Application.MacroOptions("TestSpeed", "' 测试速度和代码自动生成, test for ""adsf"" ""adsf"""" " & vbCrLf & "' " & vbCrLf & "' author: zhangzq@citics, 2009")
    Call Application.MacroOptions("PreTradingday", "' 获取前一个交易日")
    Call Application.MacroOptions("GetSecurityICODE", "' 查找对应股票或指数的代码")
    Call Application.MacroOptions("GetSecurityName", "")
    Call Application.MacroOptions("GetAccountSecurity", "")
    Call Application.MacroOptions("GetAccountIndustry", "")
    Call Application.MacroOptions("GetAccountSetupDate", "")
    Call Application.MacroOptions("GetSecurityPrice", "' 获取证券的高频价格数据")
    gRegisterUDF FunctionName:="TestSpeed1", Args:="n,s", DescriptionArgs:="""asd"",""asdf""", Category:="User Defined", Description:="test"
End Sub

Public Sub gUnregisterUDF(ByVal FunctionName As String)
    With Application
        .ExecuteExcel4Macro "UNREGISTER(" & FunctionName & ")"
        .ExecuteExcel4Macro "REGISTER(""user32.dll"" " & _
         ",""CharPrevA"",""P"",""" & FunctionName & """,,0)"
        .ExecuteExcel4Macro "UNREGISTER(" & FunctionName & ")"
    End With
End Sub

Public Sub gRegisterUDF( _
    ByVal FunctionName As String, _
    Optional ByVal Category As String, _
    Optional ByVal Description As String, _
    Optional ByVal Args As String, _
    Optional ByVal DescriptionArgs As String)
    Application.ExecuteExcel4Macro _
        "REGISTER(""user32.dll"",""CharPrevA"",""PPP"",""" _
        & FunctionName & """,""" & Args & """,1" _
        & ",""" & Category & """,,,""" & Description & """," & DescriptionArgs & ")"
End Sub

' 获取Excel实例，如果当前已经有excel窗口，则直接获取当前实例，否则新建实例
Function GetExcelApp() As Excel.Application
    On Error Resume Next
    Set GetExcelApp = GetObject(, "Excel.Application")
    If Err.Number > 0 Then Err.Clear: Set GetExcelApp = New Excel.Application
End Function
