VERSION 5.00
Begin VB.Form VB和SQL代码转化器 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   2865
   ClientTop       =   2130
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   8685
   Begin VB.CommandButton 转换成CommandText 
      Caption         =   "转换成CommandText"
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton 转化成SQL代码 
      Caption         =   "转化成SQL代码"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton 转换成VB代码 
      Caption         =   "转换成VB代码"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   5280
      Width           =   2655
   End
   Begin VB.TextBox CodeArea 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "VB和SQL代码转化器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 

Private Sub Form_Activate()
    CodeArea.Text = Clipboard.GetText
End Sub

Private Sub Form_Initialize()
    CodeArea.Text = Clipboard.GetText
    Form_Resize
End Sub



Private Sub Form_Resize()
    CodeArea.Width = Me.Width - 200
    CodeArea.Height = Me.Height - 1500
    CodeArea.Left = 50
    
    With 转换成VB代码
        .Width = 1500
        .Height = 500
        .Left = 1000
        .Top = Me.Height - 1000
    End With
    
    With 转化成SQL代码
        .Width = 1500
        .Height = 500
        .Left = 5000
        .Top = Me.Height - 1000
    End With
End Sub

Private Sub 转化成SQL代码_Click()
    Dim str As String
    str = CodeArea.Text
    
    str = Replace(str, " " & Chr(34) & " & _ " & vbCrLf & Chr(34), vbCrLf)
    str = Replace(str, """""", """")
    CodeArea.Text = str
    
    Clipboard.SetText str
End Sub


Private Sub 转换成CommandText_Click()
    Dim str As String, tmp
    str = " " & CodeArea.Text
    ' MsgBox (str)
    str = Replace(str, """", """""")
    tmp = Split(str, vbCrLf)
    str = Join(tmp, " "", "" ")
    str = "Array("" " & str & " "")"
    
    CodeArea.Text = str
    
    Clipboard.SetText str

End Sub

' 讲字符串转换成vb里可引用的字符串
Private Sub 转换成VB代码_Click()
    Dim str As String
    str = " " & CodeArea.Text
    ' MsgBox (str)
    str = Replace(str, """", """""")
    str = Replace(str, vbCrLf, " " & VBA.Chr(34) & " & _ " & vbCrLf & VBA.Chr(34)) & " "
    str = """" & str & """"
    
    CodeArea.Text = str
    
    Clipboard.SetText str

    
    ' VB和SQL代码转化器.Hide
End Sub

Private Sub Text1_Change()

End Sub
