VERSION 5.00
Begin VB.Form 选择 
   Caption         =   "查看研究报告"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.TextBox 输入证券代码 
      Height          =   270
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "选择"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   10815
      Begin VB.TextBox 起始时间 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Text            =   "2009-08-30"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox 选择预期年份 
         Height          =   300
         ItemData        =   "frm.frx":0000
         Left            =   360
         List            =   "frm.frx":0010
         TabIndex        =   3
         Text            =   "2010"
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub OLE1_Updated(Code As Integer)

End Sub


Private Sub 起始时间_Change()
    If Len(起始时间.Text) = 10 Then
        On Error GoTo endsub
        DateValue (起始时间.Text)
        RetriveYZYQData
    End If
endsub:
End Sub

Private Sub 一致预期表1_DblClick()
    MsgBox database.QueryOne("select crr.content from CMB_REPORT_RESEARCH crr " & _
            "left join i_org_info ioi on " & _
            "    ioi.id = crr.organ_id " & _
            "where crr.create_date = '" & 一致预期表1.TextMatrix(一致预期表1.MouseRow, 2) & "' " & _
            "    and ioi.org_name = '" & 一致预期表1.TextMatrix(一致预期表1.MouseRow, 3) & "' " & _
            "    and crr.author = '" & 一致预期表1.TextMatrix(一致预期表1.MouseRow, 4) & "' " & _
            "    and crr.code = '" & 一致预期表1.TextMatrix(一致预期表1.MouseRow, 0) & "';", "yzyq")
End Sub

Private Sub Form_Resize()
    一致预期表1.Width = 选择.Width - 200
End Sub

Private Sub 输入证券代码_Change()
    RetriveYZYQData
End Sub

Private Sub 选择预期年份_Click()
    RetriveYZYQData
End Sub

Private Sub 输入证券代码_GotFocus()
    输入证券代码.SelStart = 0
    输入证券代码.SelLength = Len(输入证券代码)
End Sub

Private Sub RetriveYZYQData()
    If Len(输入证券代码.Text) <> 6 Then Exit Sub
    
    一致预期表1.Clear
    
    Dim rst As New ADODB.Recordset
    Set rst = database.Query("select  " & _
            "    crr.code ""证券代码"", crr.code_name ""证券名称"",     crr.create_date ""报告日期"", " & _
            "    ioi.org_name ""报告机构"", crr.author ""作者"", " & _
            "    (case   when crr.score_id = 1 then '卖出'  " & _
            "        when crr.score_id = 2 then '派发' " & _
            "        when crr.score_id = 3 then '中性' " & _
            "        when crr.score_id = 5 then '收集' " & _
            "        when crr.score_id = 7 then '买入'  " & _
            "        else '无评级'   end) ""评级"",  " & _
            "    crr.text5 ""目标价"",    crs.time_year ""预测年份"",  " & _
            "    crs.forecast_income_share ""预测eps"" , crs.forecast_income_share/cf.c1 - 1 ""增长率"", " & _
            "    (case datalength(crr.content) when 0 then  '无' else '双击' end)  详情  " & _
            "from cmb_REPORT_RESEARCH crr   " & _
            "left join i_org_info ioi on " & _
            "    crr.organ_id = ioi.id " & _
            "left join cmb_REPORT_SUBTABLE  crs on " & _
            "    crs.report_search_id = crr.id " & _
            "left join CON_FORECAST  cf on " & _
            "    cf.stock_code = crr.Code And cf.con_date = '" & data.PreTradingday() & "'" & _
            "    and rpt_date = crs.time_year - 1 and stock_type = '1' " & _
            "where  " & _
            "    crr.create_date > '" & 起始时间.Text & "'  " & _
            "    and crs.time_year = '" & 选择预期年份.Text & "' " & _
            "    and crr.code = '" & 输入证券代码.Text & "' order by crr.create_date desc", "yzyq")
    
    Set 一致预期表1.DataSource = rst
    
    Dim i As Long
    For i = 1 To 一致预期表1.Rows - 1
        一致预期表1.TextMatrix(i, 8) = Format(一致预期表1.TextMatrix(i, 8), "0.00")
        一致预期表1.TextMatrix(i, 6) = Format(一致预期表1.TextMatrix(i, 6), "0.00")
        一致预期表1.TextMatrix(i, 9) = Format(一致预期表1.TextMatrix(i, 9), "0.0%")
    Next i
End Sub


