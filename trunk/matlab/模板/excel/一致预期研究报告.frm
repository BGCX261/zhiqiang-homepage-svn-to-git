VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ѡ�� 
   Caption         =   "�鿴�о�����"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid һ��Ԥ�ڱ�1 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10610
      _Version        =   393216
      Cols            =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      GridLines       =   3
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox ����֤ȯ���� 
      Height          =   270
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ��"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   10815
      Begin VB.TextBox ��ʼʱ�� 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Text            =   "2009-08-30"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox ѡ��Ԥ����� 
         Height          =   300
         ItemData        =   "һ��Ԥ���о�����.frx":0000
         Left            =   360
         List            =   "һ��Ԥ���о�����.frx":0010
         TabIndex        =   3
         Text            =   "2010"
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "ѡ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub OLE1_Updated(Code As Integer)

End Sub


Private Sub ��ʼʱ��_Change()
    If Len(��ʼʱ��.Text) = 10 Then
        On Error GoTo endsub
        DateValue (��ʼʱ��.Text)
        RetriveYZYQData
    End If
endsub:
End Sub

Private Sub һ��Ԥ�ڱ�1_DblClick()
    MsgBox database.QueryOne("select crr.content from CMB_REPORT_RESEARCH crr " & _
            "left join i_org_info ioi on " & _
            "    ioi.id = crr.organ_id " & _
            "where crr.create_date = '" & һ��Ԥ�ڱ�1.TextMatrix(һ��Ԥ�ڱ�1.MouseRow, 2) & "' " & _
            "    and ioi.org_name = '" & һ��Ԥ�ڱ�1.TextMatrix(һ��Ԥ�ڱ�1.MouseRow, 3) & "' " & _
            "    and crr.author = '" & һ��Ԥ�ڱ�1.TextMatrix(һ��Ԥ�ڱ�1.MouseRow, 4) & "' " & _
            "    and crr.code = '" & һ��Ԥ�ڱ�1.TextMatrix(һ��Ԥ�ڱ�1.MouseRow, 0) & "';", "yzyq")
End Sub

Private Sub Form_Resize()
    һ��Ԥ�ڱ�1.Width = ѡ��.Width - 200
End Sub

Private Sub ����֤ȯ����_Change()
    RetriveYZYQData
End Sub

Private Sub ѡ��Ԥ�����_Click()
    RetriveYZYQData
End Sub

Private Sub ����֤ȯ����_GotFocus()
    ����֤ȯ����.SelStart = 0
    ����֤ȯ����.SelLength = Len(����֤ȯ����)
End Sub

Private Sub RetriveYZYQData()
    If Len(����֤ȯ����.Text) <> 6 Then Exit Sub
    
    һ��Ԥ�ڱ�1.Clear
    
    Dim rst As New ADODB.Recordset
    Set rst = database.Query("select  " & _
            "    crr.code ""֤ȯ����"", crr.code_name ""֤ȯ����"",     crr.create_date ""��������"", " & _
            "    ioi.org_name ""�������"", crr.author ""����"", " & _
            "    (case   when crr.score_id = 1 then '����'  " & _
            "        when crr.score_id = 2 then '�ɷ�' " & _
            "        when crr.score_id = 3 then '����' " & _
            "        when crr.score_id = 5 then '�ռ�' " & _
            "        when crr.score_id = 7 then '����'  " & _
            "        else '������'   end) ""����"",  " & _
            "    crr.text5 ""Ŀ���"",    crs.time_year ""Ԥ�����"",  " & _
            "    crs.forecast_income_share ""Ԥ��eps"" , crs.forecast_income_share/cf.c1 - 1 ""������"", " & _
            "    (case datalength(crr.content) when 0 then  '��' else '˫��' end)  ����  " & _
            "from cmb_REPORT_RESEARCH crr   " & _
            "left join i_org_info ioi on " & _
            "    crr.organ_id = ioi.id " & _
            "left join cmb_REPORT_SUBTABLE  crs on " & _
            "    crs.report_search_id = crr.id " & _
            "left join CON_FORECAST  cf on " & _
            "    cf.stock_code = crr.Code And cf.con_date = '" & data.PreTradingday() & "'" & _
            "    and rpt_date = crs.time_year - 1 and stock_type = '1' " & _
            "where  " & _
            "    crr.create_date > '" & ��ʼʱ��.Text & "'  " & _
            "    and crs.time_year = '" & ѡ��Ԥ�����.Text & "' " & _
            "    and crr.code = '" & ����֤ȯ����.Text & "' order by crr.create_date desc", "yzyq")
    
    Set һ��Ԥ�ڱ�1.DataSource = rst
    
    Dim i As Long
    For i = 1 To һ��Ԥ�ڱ�1.Rows - 1
        һ��Ԥ�ڱ�1.TextMatrix(i, 8) = Format(һ��Ԥ�ڱ�1.TextMatrix(i, 8), "0.00")
        һ��Ԥ�ڱ�1.TextMatrix(i, 6) = Format(һ��Ԥ�ڱ�1.TextMatrix(i, 6), "0.00")
        һ��Ԥ�ڱ�1.TextMatrix(i, 9) = Format(һ��Ԥ�ڱ�1.TextMatrix(i, 9), "0.0%")
    Next i
End Sub


