VERSION 5.00
Begin VB.Form VB��SQL����ת���� 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   2865
   ClientTop       =   2130
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   8685
   Begin VB.CommandButton ת����CommandText 
      Caption         =   "ת����CommandText"
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton ת����SQL���� 
      Caption         =   "ת����SQL����"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CommandButton ת����VB���� 
      Caption         =   "ת����VB����"
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
Attribute VB_Name = "VB��SQL����ת����"
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
    
    With ת����VB����
        .Width = 1500
        .Height = 500
        .Left = 1000
        .Top = Me.Height - 1000
    End With
    
    With ת����SQL����
        .Width = 1500
        .Height = 500
        .Left = 5000
        .Top = Me.Height - 1000
    End With
End Sub

Private Sub ת����SQL����_Click()
    Dim str As String
    str = CodeArea.Text
    
    str = Replace(str, " " & Chr(34) & " & _ " & vbCrLf & Chr(34), vbCrLf)
    str = Replace(str, """""", """")
    CodeArea.Text = str
    
    Clipboard.SetText str
End Sub


Private Sub ת����CommandText_Click()
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

' ���ַ���ת����vb������õ��ַ���
Private Sub ת����VB����_Click()
    Dim str As String
    str = " " & CodeArea.Text
    ' MsgBox (str)
    str = Replace(str, """", """""")
    str = Replace(str, vbCrLf, " " & VBA.Chr(34) & " & _ " & vbCrLf & VBA.Chr(34)) & " "
    str = """" & str & """"
    
    CodeArea.Text = str
    
    Clipboard.SetText str

    
    ' VB��SQL����ת����.Hide
End Sub

Private Sub Text1_Change()

End Sub
