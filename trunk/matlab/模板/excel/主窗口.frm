VERSION 5.00
Begin VB.MDIForm ƽ̨ 
   BackColor       =   &H8000000C&
   Caption         =   "����֤ȯ����о�ƽ̨"
   ClientHeight    =   8715
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   12135
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu һ��Ԥ�� 
      Caption         =   "һ��Ԥ��"
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu VBת����SQL 
         Caption         =   "VBת����SQL"
      End
   End
   Begin VB.Menu ��� 
      Caption         =   "���"
      Begin VB.Menu ��ϳֲ����� 
         Caption         =   "��ϳֲ�����"
      End
      Begin VB.Menu �о�������ǽ 
         Caption         =   "�о�������ǽ"
      End
   End
End
Attribute VB_Name = "ƽ̨"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VBת����SQL_Click()
    VB��SQL����ת����.Show
End Sub

Private Sub �о�������ǽ_Click()
    �о��������ǽ.Show
End Sub

Private Sub һ��Ԥ��_Click()
    ѡ��.Show
End Sub


Private Sub ��ϳֲ�����_Click()
    �ֲ����.Show
End Sub
