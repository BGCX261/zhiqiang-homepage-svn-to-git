VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VB��SQL����ת���� 
   Caption         =   "VB��SQL����ת����"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9840
   OleObjectBlob   =   "VB��SQL����ת����.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "VB��SQL����ת����"
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



Private Sub ת����SQL����_Click()
    Dim str As String
    str = CodeArea.text
    
    str = Replace(str, vbCrLf, " " & Chr(34) & " & _ " & vbCrLf & Chr(34), vbCrLf)
    str = Replace(str, """""", """")
    CodeArea.text = str
    
    Dim MyData As New DataObject
    MyData.SetText str
    MyData.PutInClipboard
End Sub


' ���ַ���ת����vb������õ��ַ���
Private Sub ת����VB����_Click()
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
    
    VB��SQL����ת����.Hide
End Sub
