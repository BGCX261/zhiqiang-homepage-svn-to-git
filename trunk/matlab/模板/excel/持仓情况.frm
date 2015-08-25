VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form 持仓情况 
   Caption         =   "持仓情况"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   8385
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame 选择帐户框 
      Caption         =   "选择帐户"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.ComboBox 选择帐户 
         Height          =   300
         ItemData        =   "持仓情况.frx":0000
         Left            =   240
         List            =   "持仓情况.frx":0019
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker 选择日期 
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
Attribute VB_Name = "持仓情况"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    选择帐户.Text = " 交易部整体 "
    选择日期.Value = DateValue(Now)
End Sub

Private Sub Form_Resize()
    选择帐户框.Left = 100
    选择帐户框.Width = Me.Width - 500
End Sub


Private Sub 选择日期_Change()
    RefreshData
End Sub

Private Sub 选择帐户_Change()
    RefreshData
End Sub


' 获取持仓数据
Private Sub RefreshData()

End Sub

