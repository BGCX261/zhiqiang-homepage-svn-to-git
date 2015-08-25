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
    平台.Show
    ' 选择.Show
    
'
'    ' Debug.Print tools.cmd.Exist("C:\Users\tr\Desktop\模板\excel\*.xlsm")
'    Debug.Print tools.cmd.Size("C:\Users\tr\Desktop\模板\周报模板.xlsm")
'    Debug.Print tools.cmd.fileName("C:\Users\tr\Desktop\模板\周报模板.xlsm")
    ' Debug.Print tools.cmd.fileName("周报模板.xlsm")
    ' TestData
    ' TestTools
    ' TestSpeed
    ' GenerateFunList
End Sub

Sub testSaveAsPDF()
' 测试saveAsPDF函数，如果失败会给出提示
    
    Dim fileName As String
    fileName = "D:\My Documents\absdfasdf.pdf"
    abc.SaveAsPDF (fileName)
    If Dir(fileName) = "" Then MsgBox ("function SaveAsPDF doesn't work")
    Kill (fileName)
End Sub

Sub testSendMailWithOutlook()
    Set ea = GetObject(, "Excel.Application")
    tools.sendMailWithOutlook , , , Array("D:\My Documents\00.自营统计表\自营统计表20091117(程总).xls"), _
             , ea.Application.Workbooks("日报模板ver2.xlsm").Worksheets("帐户汇总").Range("B4:G15")
End Sub

Sub testsendMailWithLotus()
' 测试sendMailWithLotus函数，需要自己去邮箱检查查看邮件和附件是否正常发送

    Dim a
    Dim mailTo
    ' mailTo = Array("zhangzq@citics.com", "mathzqy@gmail.com")
    ' mailFile = Array("D:\My Documents\work\excel\zhiqiang.exp", _
            "D:\My Documents\01.风险监控日报底稿\abcd.pdf", _
            "D:\My Documents\work\excel\zhiqiang.lib")
    
    Dim EB
    Set EB = GetObject(, "Excel.Application")
    abc.sendMailWithLotus "zhangzq@citics.com;mathzqy@gmail.com", _
            "test", "hello", Array("D:\My Documents\work\excel\zhiqiang.exp", _
            "D:\My Documents\work\excel\zhiqiang.lib"), True, ea.Application.Workbooks("日报模板ver2.xlsm").Worksheets("帐户汇总").Range("B4:G15")
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
'    Set wb = ea.Workbooks("数据模块.xlsm")
'    Set ws = wb.Worksheets("基本信息")
    
    ' On Error Resume Next
    If tools.cmd.Exist("C:\DataCache.txt") Then Kill "C:\DataCache.txt"
    ' test database
    Debug.Print data.getsecurityicode("中信标普300", "指数")
    
    ' 测试收盘价
    Debug.Assert data.getsecurityinfo("山东黄金", #12/2/2009#, "收盘价") = 91.88
    Debug.Print data.getsecurityinfo("600015", , "总股本")
    Debug.Print data.getsecurityinfo("000024", "2009-11-24", "平均价")
    ' 测试获取股票代码
    Debug.Assert data.getsecurityicode("大商股份") = "600694"
    ' 测试获取股票申万行业
    Debug.Assert data.getsecurityinfo("中国银行", , "申万一级") = "金融服务"
    Debug.Assert data.getsecurityinfo("中国银行", , "中标二级") = "银行"
    ' 测试一致预期数据
    Debug.Assert data.getsecurityinfo("飞乐音响", 40151, "目标价") = "未评级"
    Debug.Assert Abs(data.getsecurityinfo("交通银行", 40151, "目标价") - 10.56) < 0.005
    
    ' 测试帐户数据
    Debug.Assert Abs(data.GetAccountInfo("专项投资", #12/4/2009#, "净值") - 32.407) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("核心组合", #12/2/2009#, "VaR") - 2.635) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("长期战略", "2009-12-02", "总成本") - 26.375) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("专项投资", "2009-12-02", "beta") - 0.92) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("专项投资", "2009-12-02", "净值") - 33.7) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("专项投资", "2009-12-02", "净值beta") - 0.76) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("配置组合", "2009-12-01", "日收益率") - 0.0021) < 0.0005
    Debug.Assert Abs(data.GetAccountInfo("中小盘", "2009-12-01", "基准收益率") - 0.0256) < 0.0005
    Debug.Assert Abs(data.GetAccountInfo("交易部整体", "2009-12-01", "日变现能力") - 0.3925) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("核心组合", "2009-12-01", "仓位增减") + 0.0061) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("交易部整体", "2009-12-01", "周变现能力") - 0.86) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("核心组合", "2009-12-01", "市值") - 30.39) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("核心组合", "2009-12-01", "单位净值") - 1.4841) < 0.0005
    Debug.Assert Abs(data.GetAccountInfo("核心组合", "2009-12-01", "止损收益率") - 0.1587) < 0.0005
    Debug.Assert data.GetAccountInfo("核心组合", , "基准") = "中信标普300"
    Debug.Assert Abs(data.GetAccountInfo("自营整体", #12/10/2009#, "现金") - 23.42) < 0.005
    Debug.Assert Abs(data.GetAccountInfo("核心组合", #12/10/2009#, "现金") - 4.322) < 0.005
    
    
    ' 测试帐户的交易数量
    Debug.Assert data.GetAccountSecurity("核心组合", "2009-11-25", "紫金矿业", "买入数量") = 5000000
    Debug.Assert data.GetAccountSecurity("核心组合", "2009-11-25", "紫金矿业", "卖出金额") = -54982409.96
    Debug.Assert Abs(data.GetAccountSecurity("核心组合", "2009-11-25", "紫金矿业", "买入价格") - 11) < 0.005
    ' 测试帐户的持仓数据
    Debug.Assert Abs(data.GetAccountSecurity("核心组合", "2009-11-25", "紫金矿业", "成本") / 10000 - 20057.99) < 0.005
    Debug.Assert Abs(data.GetAccountSecurity("核心组合", "2009-11-25", "紫金矿业", "市值") / 10000 - 21956.71) < 0.005
    Debug.Assert Abs(data.GetAccountSecurity("核心组合", "2009-11-25", "紫金矿业", "成本价") - 10.03) < 0.0001
    
    ' 测试行业数据
    Debug.Assert Abs(data.GetAccountInfo("核心组合", #12/10/2009#, "市值") - _
            data.GetAccountIndustry("核心组合", #12/10/2009#, "A股", "持有市值")) < 0.0005
    Debug.Assert data.GetAccountIndustry("交易部整体", 40151, "A股", "持有市值") > 0
    Debug.Assert data.GetAccountIndustry("配置组合", 40151, "企业债", "持有成本") = 0.3
    Debug.Assert Abs(data.GetAccountIndustry("配置组合", 40151, "企业债", "持有市值") - 0.297) < 0.0005
    Debug.Assert Abs(data.GetAccountIndustry("配置组合", #12/10/2009#, "定向增发", "持有成本") - 5.88) < 0.0005
    ' 测试帐户的绩效数据
    Debug.Assert Abs(data.ComputeDayBF("专项投资", "原材料、初级产品工业", "投资比例", 40149) - 0.831) < 0.005
    Debug.Assert Abs(data.ComputeDayBF("专项投资", "原材料、初级产品工业", "组合收益率", 40149) - 0.08) < 0.005
    Debug.Assert Abs(data.ComputeDayBF("专项投资", "原材料、初级产品工业", "基准收益率", 40149) - 0.022) < 0.005
    Debug.Assert Abs(data.ComputeDayBF("专项投资", "股票", "组合收益率", 40149) - 0.0737) < 0.005
    ' 测试止损线
    
    ' 测试存储过程
    Dim rst As ADODB.Recordset
    Set rst = database.QueryStoredProc("REP_ZHIQIANG_RETURN_RISK_SY", Array(3201, 30626, #12/4/2009#, "1", 95, 1, "2010"))
    Debug.Assert rst.Fields.Count = 14
    
    
    
End Sub

Private Sub TestTools()
    '
    ' tools.cmd.Open_ "C:\Users\tr\Desktop\模板\eula.txt"
    
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
