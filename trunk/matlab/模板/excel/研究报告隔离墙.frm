VERSION 5.00
Begin VB.Form �о��������ǽ 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   9420
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ȷ�� 
      Caption         =   "��ѯ�ֲ�"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox �о��������� 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "�о��������ǽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    �о���������.Left = 100
    �о���������.Width = Me.Width - 500
End Sub



Private Sub RefreshData(SecuList$)

End Sub

Private Sub ȷ��_Click()
    Dim t$, SecuList$, i&, ind&
    t = �о���������.Text
    i = 1
    ind = 0
    Do While i <= Len(t)
        Dim tmp$
        tmp = VBA.Mid$(t, i, 1)
        If tmp <= "9" And tmp >= "0" Then
            ind = ind + 1
            If ind = 6 Then
                If SecuList = "" Then
                    SecuList = "'" & VBA.Mid(t, i - 5, 6) & "'"
                Else
                    SecuList = SecuList & ", '" & VBA.Mid(t, i - 5, 6) & "'"
                End If
            End If
        Else
            ind = 0
        End If
        
        i = i + 1
    Loop
    
    MsgBox SecuList
End Sub
