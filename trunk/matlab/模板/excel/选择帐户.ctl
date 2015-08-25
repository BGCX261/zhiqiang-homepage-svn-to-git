VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl 选择帐户 
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   ScaleHeight     =   1680
   ScaleWidth      =   6120
   Begin VB.Frame 选择帐户框 
      Caption         =   "选择帐户"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   50200577
         CurrentDate     =   40144
      End
      Begin VB.ComboBox 选择帐户代码 
         Height          =   300
         ItemData        =   "选择帐户.ctx":0000
         Left            =   240
         List            =   "选择帐户.ctx":0019
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   480
         Width           =   2655
      End
   End
End
Attribute VB_Name = "选择帐户"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub
