VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VB和SQL代码转换器 
   Caption         =   "VB和SQL代码转换器"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9840
   OleObjectBlob   =   "VB和SQL代码转换器.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "VB和SQL代码转换器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CodeArea_Change()

End Sub

Private Sub UserForm_Activate()
    Dim MyData As New DataObject
    MyData.GetFromClipboard
    CodeArea.text = MyData.GetText(1)
End Sub

Private Sub UserForm_Initialize()
    Dim MyData As New DataObject
    MyData.GetFromClipboard
    CodeArea.text = MyData.GetText(1)
End Sub



Private Sub 转化成SQL代码_Click()
    Dim str As String
    str = CodeArea.text
    
    str = Replace(str, vbCrLf, " " & Chr(34) & " & _ " & vbCrLf & Chr(34), vbCrLf)
    str = Replace(str, """""", """")
    CodeArea.text = str
    
    Dim MyData As New DataObject
    MyData.SetText str
    MyData.PutInClipboard
End Sub


' 讲字符串转换成vb里可引用的字符串
Private Sub 转换成VB代码_Click()
    Dim str As String
    str = CodeArea.text
    ' MsgBox (str)
    str = Replace(str, """", """""")
    str = Replace(str, vbCrLf, " " & Chr(34) & " & _ " & vbCrLf & Chr(34))
    str = """" & str & """"
    
    CodeArea.text = str
    
    Dim MyData As New DataObject
    MyData.SetText str
    MyData.PutInClipboard
    
    VB和SQL代码转换器.Hide
End Sub
