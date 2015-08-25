VERSION 5.00
Begin VB.MDIForm 平台 
   BackColor       =   &H8000000C&
   Caption         =   "中信证券风控研究平台"
   ClientHeight    =   8715
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   12135
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu 一致预期 
      Caption         =   "一致预期"
   End
   Begin VB.Menu 工具 
      Caption         =   "工具"
      Begin VB.Menu VB转化成SQL 
         Caption         =   "VB转化成SQL"
      End
   End
   Begin VB.Menu 组合 
      Caption         =   "组合"
      Begin VB.Menu 组合持仓详情 
         Caption         =   "组合持仓详情"
      End
      Begin VB.Menu 研究部隔离墙 
         Caption         =   "研究部隔离墙"
      End
   End
End
Attribute VB_Name = "平台"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VB转化成SQL_Click()
    VB和SQL代码转化器.Show
End Sub

Private Sub 研究部隔离墙_Click()
    研究报告隔离墙.Show
End Sub

Private Sub 一致预期_Click()
    选择.Show
End Sub


Private Sub 组合持仓详情_Click()
    持仓情况.Show
End Sub
