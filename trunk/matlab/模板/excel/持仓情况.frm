VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form �ֲ���� 
   Caption         =   "�ֲ����"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   8385
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame ѡ���ʻ��� 
      Caption         =   "ѡ���ʻ�"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.ComboBox ѡ���ʻ� 
         Height          =   300
         ItemData        =   "�ֲ����.frx":0000
         Left            =   240
         List            =   "�ֲ����.frx":0019
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker ѡ������ 
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   130940929
         CurrentDate     =   40144
      End
   End
End
Attribute VB_Name = "�ֲ����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    ѡ���ʻ�.Text = " ���ײ����� "
    ѡ������.Value = DateValue(Now)
End Sub

Private Sub Form_Resize()
    ѡ���ʻ���.Left = 100
    ѡ���ʻ���.Width = Me.Width - 500
End Sub


Private Sub ѡ������_Change()
    RefreshData
End Sub

Private Sub ѡ���ʻ�_Change()
    RefreshData
End Sub


' ��ȡ�ֲ�����
Private Sub RefreshData()

End Sub

