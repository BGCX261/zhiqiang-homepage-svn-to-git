Attribute VB_Name = "testMain"
Public abc As New CITICSTools
Public tools As New CITICSTools
Public data As New CITICSData
Public database As New CITICSDatabase
Public ea As Excel.Application

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_ACTIVATE As Long = &H6

Sub Main()
    ' testSaveAsPDF
    ' testsendMailWithLotus
    ' testSendNotesMail
    
    ' Debug.Print "main test start"
    ' Set ea = GetObject(, "Excel.Application")
    ' data.test
    
    ' testSaveAsPDF
    ' testsendMailWithLotus
    ' testSendMailWithOutlook
    ƽ̨.Show
    ' ѡ��.Show
    
'
'    ' Debug.Print tools.cmd.Exist("C:\Users\tr\Desktop\ģ��\excel\*.xlsm")
'    Debug.Print tools.cmd.Size("C:\Users\tr\Desktop\ģ��\�ܱ�ģ��.xlsm")
'    Debug.Print tools.cmd.fileName("C:\Users\tr\Desktop\ģ��\�ܱ�ģ��.xlsm")
    ' Debug.Print tools.cmd.fileName("�ܱ�ģ��.xlsm")
    ' TestData
    ' TestTools
    ' TestSpeed
    ' GenerateFunList
End Sub

Sub testSaveAsPDF()
' ����saveAsPDF���������ʧ�ܻ������ʾ
    
    Dim fileName As String
    fileName = "D:\My Documents\absdfasdf.pdf"
    abc.SaveAsPDF (fileName)
    If Dir(fileName) = "" Then MsgBox ("function SaveAsPDF doesn't work")
    Kill (fileName)
End Sub

Sub testSendMailWithOutlook()
    Set ea = GetObject(, "Excel.Application")
    tools.sendMailWithOutlook , , , Array("D:\My Documents\00.��Ӫͳ�Ʊ�\��Ӫͳ�Ʊ�20091117(����).xls"), _
             , ea.Application.Workbooks("�ձ�ģ��ver2.xlsm").Worksheets("�ʻ�����").Range("B4:G15")
End Sub

Sub testsendMailWithLotus()
' ����sendMailWithLotus��������Ҫ�Լ�ȥ������鿴�ʼ��͸����Ƿ���������

    Dim a
    Dim mailTo
    ' mailTo = Array("zhangzq@citics.com", "mathzqy@gmail.com")
    ' mailFile = Array("D:\My Documents\work\excel\zhiqiang.exp", _
            "D:\My Documents\01.���ռ���ձ��׸�\abcd.pdf", _
            "D:\My Documents\work\excel\zhiqiang.lib")
    
    Dim EB
    Set EB = GetObject(, "Excel.Application")
    abc.sendMailWithLotus "zhangzq@citics.com;mathzqy@gmail.com", _
            "test", "hello", Array("D:\My Documents\work\excel\zhiqiang.exp", _
            "D:\My Documents\work\excel\zhiqiang.lib"), True, ea.Application.Workbooks("�ձ�ģ��ver2.xlsm").Worksheets("�ʻ�����").Range("B4:G15")
End Sub


Private Sub ClearDebugMessage()
    Dim ideHwnd&, debugFrmHwnd&
    ideHwnd = FindWindow("wndclass_desked_gsk", vbNullString)
    If ideHwnd > 0 Then
            debugFrmHwnd = FindWindowEx(ideHwnd, ByVal 0&, "VbaWindow", vbNullString)
            If debugFrmHwnd > 0 Then
                PostMessage debugFrmHwnd, WM_ACTIVATE, 1, 0&
                SendKeys "^{HOME}+^{END}^{BREAK}{DEL}{F5}", False
            End If
    End If
End Sub

Private Sub TestData()
    Dim ea As Excel.Application, ws As Excel.Worksheet, wb As Excel.Workbook
'    Set ea = GetObject(, "Excel.Application")
'    Set wb = ea.Workbooks("����ģ��.xlsm")
'    Set ws = wb.Worksheets("������Ϣ")
    
    ' On Error Resume Next
    If tools.cmd.Exist("C:\DataCache.txt") Then Kill "C:\DataCache.txt"
    ' test database
    Debug.Print data.getsecurityicode("���ű���300", "ָ��")
    
    ' �������̼�
    Debug.Assert data.getsecurityinfo("ɽ���ƽ�", #12/2/2009#, "���̼�") = 91.88
    Debug.Print data.getsecurityinfo("600015", , "�ܹɱ�")
    Debug.Print data.getsecurityinfo("000024", "2009-11-24", "ƽ����")
    ' ���Ի�ȡ��Ʊ����
    Debug.Assert data.getsecurityicode("���̹ɷ�") = "600694"
    ' ���Ի�ȡ��Ʊ������ҵ
    Debug.Assert data.getsecurityinfo("�й�����", , "����һ��") = "���ڷ���"
    Debug.Assert data.getsecurityinfo("�й�����", , "�б����") = "����"
    ' ����һ��Ԥ������
    Debug.Assert data.getsecurityinfo("��������", 40151, "Ŀ���") = "δ����"
    Debug.Assert Abs(data.getsecurityinfo("��ͨ����", 40151, "Ŀ���") - 10.56) < 0.005
    
    ' �����ʻ�����
    Debug.Assert Abs(data.GetAccountInfo("ר��Ͷ��", #12/4/2009#, "��ֵ") - 32.407) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("�������", #12/2/2009#, "VaR") - 2.635) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("����ս��", "2009-12-02", "�ܳɱ�") - 26.375) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("ר��Ͷ��", "2009-12-02", "beta") - 0.92) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("ר��Ͷ��", "2009-12-02", "��ֵ") - 33.7) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("ר��Ͷ��", "2009-12-02", "��ֵbeta") - 0.76) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("�������", "2009-12-01", "��������") - 0.0021) < 0.0005
    Debug.Assert Abs(data.GetAccountInfo("��С��", "2009-12-01", "��׼������") - 0.0256) < 0.0005
    Debug.Assert Abs(data.GetAccountInfo("���ײ�����", "2009-12-01", "�ձ�������") - 0.3925) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("�������", "2009-12-01", "��λ����") + 0.0061) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("���ײ�����", "2009-12-01", "�ܱ�������") - 0.86) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("�������", "2009-12-01", "��ֵ") - 30.39) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("�������", "2009-12-01", "��λ��ֵ") - 1.4841) < 0.0005
    Debug.Assert Abs(data.GetAccountInfo("�������", "2009-12-01", "ֹ��������") - 0.1587) < 0.0005
    Debug.Assert data.GetAccountInfo("�������", , "��׼") = "���ű���300"
    Debug.Assert Abs(data.GetAccountInfo("��Ӫ����", #12/10/2009#, "�ֽ�") - 23.42) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("�������", #12/10/2009#, "�ֽ�") - 4.322) < 0.005
    
    
    ' �����ʻ��Ľ�������
    Debug.Assert data.GetAccountSecurity("�������", "2009-11-25", "�Ͻ��ҵ", "��������") = 5000000
    Debug.Assert data.GetAccountSecurity("�������", "2009-11-25", "�Ͻ��ҵ", "�������") = -54982409.96
    Debug.Assert Abs(data.GetAccountSecurity("�������", "2009-11-25", "�Ͻ��ҵ", "����۸�") - 11) < 0.005
    ' �����ʻ��ĳֲ�����
    Debug.Assert Abs(data.GetAccountSecurity("�������", "2009-11-25", "�Ͻ��ҵ", "�ɱ�") / 10000 - 20057.99) < 0.005
    Debug.Assert Abs(data.GetAccountSecurity("�������", "2009-11-25", "�Ͻ��ҵ", "��ֵ") / 10000 - 21956.71) < 0.005
    Debug.Assert Abs(data.GetAccountSecurity("�������", "2009-11-25", "�Ͻ��ҵ", "�ɱ���") - 10.03) < 0.0001
    
    ' ������ҵ����
    Debug.Assert Abs(data.GetAccountInfo("�������", #12/10/2009#, "��ֵ") - _
            data.GetAccountIndustry("�������", #12/10/2009#, "A��", "������ֵ")) < 0.0005
    Debug.Assert data.GetAccountIndustry("���ײ�����", 40151, "A��", "������ֵ") > 0
    Debug.Assert data.GetAccountIndustry("�������", 40151, "��ҵծ", "���гɱ�") = 0.3
    Debug.Assert Abs(data.GetAccountIndustry("�������", 40151, "��ҵծ", "������ֵ") - 0.297) < 0.0005
    Debug.Assert Abs(data.GetAccountIndustry("�������", #12/10/2009#, "��������", "���гɱ�") - 5.88) < 0.0005
    ' �����ʻ��ļ�Ч����
    Debug.Assert Abs(data.ComputeDayBF("ר��Ͷ��", "ԭ���ϡ�������Ʒ��ҵ", "Ͷ�ʱ���", 40149) - 0.831) < 0.005
    Debug.Assert Abs(data.ComputeDayBF("ר��Ͷ��", "ԭ���ϡ�������Ʒ��ҵ", "���������", 40149) - 0.08) < 0.005
    Debug.Assert Abs(data.ComputeDayBF("ר��Ͷ��", "ԭ���ϡ�������Ʒ��ҵ", "��׼������", 40149) - 0.022) < 0.005
    Debug.Assert Abs(data.ComputeDayBF("ר��Ͷ��", "��Ʊ", "���������", 40149) - 0.0737) < 0.005
    ' ����ֹ����
    
    ' ���Դ洢����
    Dim rst As ADODB.Recordset
    Set rst = database.QueryStoredProc("REP_ZHIQIANG_RETURN_RISK_SY", Array(3201, 30626, #12/4/2009#, "1", 95, 1, "2010"))
    Debug.Assert rst.Fields.Count = 14
    
    
    
End Sub

Private Sub TestTools()
    '
    ' tools.cmd.Open_ "C:\Users\tr\Desktop\ģ��\eula.txt"
    
    Debug.Assert tools.ColumnLetter(tools.ColumnNumber("A")) = "A"
    Debug.Assert tools.ColumnLetter(tools.ColumnNumber("BA")) = "BA"
    Debug.Assert tools.ColumnLetter(tools.ColumnNumber("ZD")) = "ZD"
    'Debug.Assert tools.ColumnLetter(tools.ColumnNumber("DEF")) = "DEF"
    Debug.Assert tools.ColumnNumber("D") = 4
    Debug.Assert tools.ColumnLetter(4) = "D"
    Debug.Assert tools.ColumnLetter(52) = "AZ"
    
    ' tools.StopRefresh True
    ' tools.StopRefresh False
End Sub

Private Sub TestSpeed()
    Dim i&, start As Double
    start = Timer
    Debug.Print data.TestSpeed(1)
    For i = 1 To 100000
        data.TestSpeed i
    Next i
    Debug.Print Timer - start
End Sub
