VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CITICSData"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' data

Public cache As New CITICSCache

Private Sub Class_Initialize()
    cache.CacheFile = "C:\DataCache.txt"
End Sub

' 设置缓存文件地址
Public Property Let CacheFile(cf$)
    cache.CacheFile = cf
End Property

' 测试速度和代码自动生成
'
' Author: zhangzq@citics, 2009
Public Function TestSpeed&(n& _
        , Optional s = "asd""")
    TestSpeed = cache.Item("TestSpeed" & n)
    If TestSpeed <> cache.NotCached Then Exit Function
    
    TestSpeed = n - 1
    
    cache.Add "TestSpeed" & n, TestSpeed
End Function

' 获取前一个交易日
' intervalDay 间隔日期，默认为1，返回前一交易日
' Author: zhangzq@citics, 2009
Public Function PreTradingday(Optional dt As Date = 0, Optional intervalDay As Long = 1)
    If dt = 0 Then dt = VBA.Date
    PreTradingday = cache.Item("前交易日" & dt & intervalDay)
    If PreTradingday <> cache.NotCached Then Exit Function
    
    If dt = 0 Then d = Date
    If intervalDay = 1 Then
        PreTradingday = database.QueryOne("select top 1 date from qt_tradingday where iftradingday = 1 " & _
            " and secumarket = 83 and  date < '" & Format(dt, "yyyy-mm-dd") & "'" & _
            " order by date desc ", "JYDB")
    Else
        PreTradingday = database.QueryOne("select top 1 date from ( " & _
                " select top " & intervalDay & " date from qt_tradingday " & _
                " where iftradingday = 1 and secumarket = 83 and  date < '" & Format(dt, "yyyy-mm-dd") & "'" & _
                " order by date desc) a order by date", "JYDB")
    End If
    
    cache.Add "前交易日" & dt & intervalDay, PreTradingday
End Function

' 获取后一个交易日
' Author: zhangzq@citics, 2009
Public Function NextTradingDay(Optional ByVal dt As Date = 0, Optional intervalDay As Long = 1)
    If dt = 0 Then d = VBA.Date
    NextTradingDay = cache.Item("后交易日" & dt & intervalDay)
    If NextTradingDay <> cache.NotCached Then Exit Function
    
    If intervalDay = 1 Then
        NextTradingDay = database.QueryOne("select top 1 date from qt_tradingday where iftradingday = 1 " & _
            " and secumarket = 83 and  date > '" & Format(dt, "yyyy-mm-dd") & "'" & _
            " order by date ", "yjdb")
    Else
        NextTradingDay = database.QueryOne("select top 1 date from ( " & _
                " select top " & intervalDay & " date from qt_tradingday " & _
                " where iftradingday = 1 and secumarket = 83 and  date > '" & Format(dt, "yyyy-mm-dd") & "'" & _
                " order by date) a order by date DESC", "jydb")
    End If
    cache.Add "后交易日" & dt & intervalDay, NextTradingDay
End Function


' 获取指数信息
' info: 收盘价、日收益率等
' Author: zhangzq@citics, 2009
Public Function GetIndexInfo(ByVal indexName$, infoDate As Date, Optional info$ = "收盘价")
    
    If indexName = "" Then Exit Function
    
    ' 先读取缓存，如果缓存里有数据，则直接返回缓存
    GetIndexInfo = cache.Item("SZ" & indexName & infoDate & info)
    If GetIndexInfo <> cache.NotCached Then Exit Function
        
    Select Case LCase(indexName)
        ' 复合基准的计算比较麻烦
        Case "中小盘复合基准"
            Select Case LCase(info)
            Case "收盘价", "close"
                GetIndexInfo = 200 * GetIndexInfo("沪深300", infoDate) _
                        / GetIndexInfo("沪深300", GetAccountSetupDate("中小盘")) + 800 * _
                        GetIndexInfo("中信小盘", infoDate) _
                        / GetIndexInfo("中信小盘", GetAccountSetupDate("中小盘"))
            Case "日收益率"
                GetIndexInfo = 0.2 * GetIndexInfo("沪深300", infoDate, "日收益率") + _
                        0.8 * GetIndexInfo("中信小盘", infoDate, "日收益率")
            Case Else
                GetIndexInfo = "#N/A"
            End Select
        ' 对于常用的几个指数信息可以从数据透视表里获取
        Case Else
            Select Case LCase(info)
                Case "收盘价", "close"
                    Dim sql As String, icode As String
                    icode = GetSecurityICODE(indexName, "指数")
                    
                    sql = "SELECT i.ClosePrice FROM QT_IndexQuote i  " & _
                            "inner join SecuMain m on m.InnerCode = i.InnerCode " & _
                            "where i.TradingDay = '" & Format(infoDate, "yyyy-mm-dd") & "' " & _
                            "   and m.SecuCode = '" & icode & "' " & _
                            "   and m.SecuCategory = '4' "
                    GetIndexInfo = database.QueryOne(sql, "JYDB")
                Case "日收益率"
                    GetIndexInfo = GetIndexInfo(indexName, infoDate, "close") / _
                            GetIndexInfo(indexName, PreTradingday(infoDate), "close") - 1
                Case Else
                    GetIndexInfo = "#N/A"
            End Select
    End Select

    cache.Add "SZ" & indexName & infoDate & info, GetIndexInfo
End Function

' 获取个股信息
' iCode: 个股代码或证券名称
' infoDate: 日期，可选
' info: 可选收盘价等行情数据，PE/PB等财务数据，总股本等公司信息，申万一级、中标一级等行业信息，一致预期信息
' Author: zhangzq@citics, 2009
Public Function GetSecurityInfo(icode_ As String, Optional infoDate As Date, Optional info$ = "收盘价")
    If icode_ = "" Then Exit Function
    
    Dim icode As String
    GetSecurityInfo = cache.Item("GP" & icode_ & infoDate & info)
    If GetSecurityInfo <> cache.NotCached Then Exit Function
        
    ' 处理输入信息，icode可以是股票代码，也可以是股票名称
    ' 如果是股票名称，则先去取回股票代码
    icode = GetSecurityICODE(icode_)
    
    Dim sql$, columnName$
        
        
    ' 获取证券的行业数据
    Dim classMethod$
    Select Case LCase(info)
        Case "日收益率"
            GetSecurityInfo = GetSecurityInfo(icode, infoDate, "收盘价") _
                    / GetSecurityInfo(icode, PreTradingday(infoDate), "收盘价") - 1
        Case "申万一级行业", "申万一级"
            classMethod = "申万研究所"
            columnName = "FirstIndustryName"
        Case "申万二级行业", "申万二级"
            classMethod = "申万研究所"
            columnName = "SecondIndustryName"
        Case "申万三级行业", "申万三级"
            classMethod = "申万研究所"
            columnName = "ThirdIndustryName"
        Case "中标一级行业", "中标一级"
            classMethod = "中信标普"
            columnName = "FirstIndustryName"
        Case "中标二级行业", "中标二级"
            classMethod = "中信标普"
            columnName = "SecondIndustryName"
        Case "中标三级行业", "中标三级"
            classMethod = "中信标普"
            columnName = "ThirdIndustryName"
    End Select
    
    If Len(columnName) Then
        sql = "select le." & columnName & " from LC_ExgIndustry le " & _
                "left join secumain sm on sm.companycode = le.companycode " & _
                "where le.infoSource = '" & classMethod & "' " & _
                "and le.CancelDate is null " & _
                "and sm.secucode='" & icode & "' " & _
                "and sm.SecuCategory <> '4'"
        GetSecurityInfo = database.QueryOne(sql, "JYDB")
        columnName = ""
    End If
    
    Select Case LCase(info)
        Case "总股本"
            columnName = "totalshares"
    End Select
    
    If Len(columnName) Then
        sql = "select " & columnName & " from LC_NewestShareStru ln " & _
                "left join secumain sm on sm.companycode = ln.companycode " & _
                "where sm.secucode = '" & icode & "'"
        GetSecurityInfo = database.QueryOne(sql, "JYDB")
        columnName = ""
    End If
    
    ' 从聚源获取证券的行情数据
    Select Case LCase(info)
        Case "收盘价", "closeprice", "close"
            columnName = "ClosePrice"
        Case "前收盘价"
            columnName = "PrevClosePrice"
        Case "最高价"
            columnName = "HighPrice"
        Case "最低价"
            columnName = "LowPrice"
        Case "开盘价"
            columnName = "OpenPrice"
        Case "交易量"
            columnName = "TurnoverVolume"
        Case "交易额", "交易金额"
            columnName = "TurnoverValue"
        Case "平均价", "交易均价"
            columnName = "TurnoverValue/(qd.TurnoverVolume+0.001)"
    End Select
    
    If Len(columnName) Then
        sql = "select qd." & columnName & " from QT_DailyQuote qd " & _
                "left join secumain sm on qd.innercode = sm.innercode " & _
                "where qd.Tradingday = '" & infoDate & "' " & _
                "and sm.secucode='" & icode & "' " & _
                "and sm.SecuCategory <> '4'"
        GetSecurityInfo = database.QueryOne(sql, "JYDB")
        columnName = ""
    End If
    
    Select Case LCase(info)
        Case "市盈率", "pe"
            columnName = "pe"
        Case "市净率", "pb"
            columnName = "pb"
        Case "eps", "roe"
            columnName = info
        Case Else
            columnName = ""
    End Select
    
    If Len(columnName) > 0 Then
        sql = "select " & columnName & " from xrisk.TBSI_ANTE_STOCK " & _
                " where I_CODE = '" & icode & "' and beg_date = '" & _
                Format(infoDate, "yyyy-mm-dd") & "'"
        GetSecurityInfo = database.QueryOne(sql, "xRisk")
        columnName = ""
    End If
    
    ' 从一致预期数据库获取盈利分析
    Dim year As String
    Select Case LCase(info)
        Case "2009一致预期pe", "2010一致预期pe", "2011一致预期pe"
            columnName = "c5"
            year = Left(info, 4)
        Case "2009一致预期pb", "2010一致预期pb", "2011一致预期pb"
            columnName = "cPB"
            year = Left(info, 4)
        Case "2009一致预期eps", "2010一致预期eps", "2011一致预期eps"
            columnName = "c1"
            year = Left(info, 4)
        Case "2009一致预期周变化率", "2010一致预期周变化率", "2011一致预期周变化率"
            columnName = "c80"
            year = Left(info, 4)
        Case "2009一致预期月变化率", "2010一致预期月变化率", "2011一致预期月变化率"
            columnName = "c81"
            year = Left(info, 4)
    End Select
    
    If Len(columnName) > 0 Then
        sql = "select " & columnName & " from CON_FORECAST  where stock_code = '" & icode & "' " _
                & " and con_date = '" & Format(infoDate, "yyyy-mm-dd") & "' " _
                & " and rpt_date = '" & year & "' and stock_type = '1'"
        GetSecurityInfo = database.QueryOne(sql, "YZYQ")
        
        If info Like "*变化率*" Then GetSecurityInfo = GetSecurityInfo / 100
        
        columnName = ""
    End If
    
    ' 从一致预期数据库获取一致预期的评级和目标价
    Select Case LCase(info)
        Case "一致预期评级", "评级"
            columnName = "score"
        Case "一致预期目标价", "目标价"
            columnName = "target_price"
    End Select
    If Len(columnName) > 0 Then
        sql = "select " & columnName & " from CON_FORECAST_SCHEDULE where stock_code='" & icode & _
                "' and con_date = '" & Format(infoDate, "yyyy-mm-dd") & "'"
        GetSecurityInfo = database.QueryOne(sql, "YZYQ")
        columnName = ""
        If GetSecurityInfo < 0 Then
            GetSecurityInfo = "未评级"
        End If
    End If
    
    cache.Add "GP" & icode_ & infoDate & info, GetSecurityInfo
End Function

' 注意时间区间为左开右闭区间
Public Function GetReturnRate(icode As String, beginDate As Date, endDate As Date)
    Select Case icode
        Case "中小盘复合基准"
            GetReturnRate = 0.2 * GetReturnRate("沪深300", beginDate, endDate) + _
                    0.8 * GetReturnRate("中信小盘", beginDate, endDate)
        Case Else
            GetReturnRate = GetIndexInfo(icode, endDate, "close") / _
                    GetIndexInfo(icode, beginDate, "close") - 1
    End Select
End Function


' 查找对应股票或指数的代码
Public Function GetSecurityICODE(ByRef name As String, Optional secuType$ = "股票")
    If Not val(mid(name, 1, 6)) = 0 Then
        GetSecurityICODE = mid(name, 1, 6)
        Exit Function
    End If
    
    GetSecurityICODE = cache.Item(secuType & name & "icode")
    If GetSecurityICODE <> cache.NotCached Then Exit Function
    
    Dim sql As String
    If secuType = "股票" Then
        sql = "select I_CODE from xrisk.vecd_main where S_NAME = '" & name & "'"
        GetSecurityICODE = database.QueryOne(sql, "xRisk")
    Else
        sql = "select SecuCode from SecuMain where SecuAbbr like '%" & name & "%'"
        GetSecurityICODE = database.QueryOne(sql, "JYDB")
    End If
    
    cache.Add secuType & name & "icode", GetSecurityICODE
End Function

' 已知股票代码查找证券名称
Public Function GetSecurityName(ByRef icode$, Optional secuType$ = "股票") As String
    If Len(icode) < 6 Then
        GetSecurityName = icode
        Exit Function
    End If
    
    GetSecurityName = cache.Item(secuType & name & "name")
    If GetSecurityName <> cache.NotCached Then Exit Function
    
    Dim sql As String
    If secuType = "股票" Then
        sql = "select S_NAME  from xrisk.vecd_main where I_CODE = '" & name & "'"
        GetSecurityName = database.QueryOne(sql, "JYDB")
    Else
        sql = "select SecuAbbr from SecuMain where SecuCode like '%" & name & "%'"
        GetSecurityName = database.QueryOne(sql, "JYDB")
    End If
        
    cache.Add secuType & name & "name", GetSecurityName
End Function

' 获取帐户信息
' account: 帐户名称
' info: 市值、成本、日收益率等等
' Author: zhangzq@citics, 2009
Public Function GetAccountInfo(account$, Optional infoDate As Date = 0, Optional info$ = "净值")
        
    If account = "" Then Exit Function
    
    GetAccountInfo = cache.Item(account & infoDate & info)
    If GetAccountInfo <> cache.NotCached Then Exit Function
    
    Select Case LCase(info)
        Case "port code", "port_code"
            GetAccountInfo = GetAccountPortCode(account)
        Case "p_id", "pid"
            GetAccountInfo = GetAccountPID(account)
    End Select
    
    Dim sql As String, column As String
    Select Case LCase(info)
        Case "基准"
            GetAccountInfo = GetAccountBase(account)
        Case "额度"
            pID = GetAccountInfo(account, , "pid")
            sql = "select fundlimit/100000000  " & _
                    "from xrisk.tmp_tfundlimit " & _
                    "where " & _
                    "    P_ID = '" & pID & "' " & _
                    "    and beg_date <= '" & Format(infoDate, "yyyy-mm-dd") & "' " & _
                    "    and end_date > '" & Format(infoDate, "yyyy-mm-dd") & "'"
            GetAccountInfo = database.QueryOne(sql, "xRisk")
        Case "仓位"
            GetAccountInfo = GetAccountInfo(account, infoDate, "市值") _
                    / GetAccountInfo(account, infoDate, "净值")
        Case "仓位变化", "仓位变动", "仓位增减"
            GetAccountInfo = GetAccountInfo(account, infoDate, "仓位") - _
                    GetAccountInfo(account, PreTradingday(infoDate), "仓位")
        Case "净值", "单位净值"
            column = IIf(info = "净值", " p_totalnav/100000000 ", " p_unitNav ")
            sql = "select " & column & _
                    "from xrisk.TCRP_NAV_ZX " & _
                    "where " & _
                    "    Port_Code = '" & GetAccountInfo(account, , "port code") & "' " & _
                    "    and beg_date ='" & Format(infoDate, "yyyy-mm-dd") & "' "
            If account = "交易部整体" Then
                GetAccountInfo = GetAccountInfo("长期战略", infoDate, "净值") _
                    + GetAccountInfo("核心组合", infoDate, "净值") _
                    + GetAccountInfo("中小盘", infoDate, "净值") _
                    + GetAccountInfo("配置组合", infoDate, "净值")
            Else
                GetAccountInfo = database.QueryOne(sql, "xRisk")
            End If
        Case "总盈亏"
            GetAccountInfo = GetAccountInfo(account, infoDate, "净值") - _
                    GetAccountInfo(account, infoDate, "额度")
        Case "浮动盈亏"
            GetAccountInfo = GetAccountInfo(account, infoDate, "市值") - _
                    GetAccountInfo(account, infoDate, "总成本")
        Case "已实现盈利"
            GetAccountInfo = GetAccountInfo(account, infoDate, "总盈亏") - _
                    GetAccountInfo(account, infoDate, "浮动盈亏")
        Case "现金", "cash"
            Select Case account
                Case "长期战略", "中小盘", "配置组合"
                    GetAccountInfo = GetAccountInfo(account, infoDate, "额度") _
                            - GetAccountInfo(account, infoDate, "总成本")
                Case "核心组合", "专项投资"
                    GetAccountInfo = GetAccountInfo(account, infoDate, "净值") _
                            - GetAccountInfo(account, infoDate, "市值")
                Case "交易部整体"
                    GetAccountInfo = GetAccountInfo("核心组合", infoDate, info) _
                            + GetAccountInfo("长期战略", infoDate, info) _
                            + GetAccountInfo("配置组合", infoDate, info) _
                            + GetAccountInfo("中小盘", infoDate, info)
                Case "自营整体"
                     GetAccountInfo = GetAccountInfo("交易部整体", infoDate, info) _
                            + GetAccountInfo("专项投资", infoDate, info)
            End Select
        Case "基准日资产"
            GetAccountInfo = GetAccountBaseNV(account)
        Case "净值beta"
            GetAccountInfo = GetAccountInfo(account, infoDate, "仓位") * _
                    GetAccountInfo(account, infoDate, "Beta")
        Case "var占比"
            GetAccountInfo = GetAccountInfo(account, infoDate, "var") / _
                    GetAccountInfo(account, infoDate, "净值")
        Case "日收益率", "当日收益率"
            GetAccountInfo = GetAccountInfo(account, infoDate, "单位净值") / _
                    GetAccountInfo(account, PreTradingday(infoDate), "单位净值") - 1
        Case "基准日收益率", "日基准收益率", "当日基准收益率", "当日基准收益"
            GetAccountInfo = GetIndexInfo(GetAccountInfo(account, , "基准"), infoDate, "日收益率")
        Case "基准累计收益"
            GetAccountInfo = GetReturnRate(GetAccountInfo(account, , "基准"), _
                    GetAccountInfo(account, , "设立时间"), infoDate)
            
        Case "日变现能力"
            sql = "select trunc(" & _
                    "    sum((case when cash_avg_day > 0 and (pd.i_code <> '600383' and pd.i_code <> '000012') " & _
                    "        then (case when (1/cash_avg_day)>1   " & _
                    "             then  1  else (1/cash_avg_day) " & _
                    "             end) " & _
                    "        else 0 end)*pd.mtm) " & _
                    "    /sum(pd.mtm) ,5) " & _
                    "from xrisk.tbsi_ante_port t " & _
                    "left join xrisk.tbsi_ante_portdetail pd on " & _
                    "    pd.pi_id=t.pi_id  " & _
                    "    and pd.w<>0 " & _
                    "    and pd.a_type like '%S%' " & _
                    "where " & _
                    "    t.p_id = '" & GetAccountPID(account) & "' " & _
                    "    and t.bm_id in ('30626', '30724') " & _
                    "    and t.beg_date = '" & Format(infoDate, "yyyy-mm-dd") & "'"
            GetAccountInfo = database.QueryOne(sql, "xRisk")
        Case "周变现能力"
            sql = "select trunc( " & _
                    "    sum((case when cash_avg_day > 0 and (pd.i_code <> '600383' and pd.i_code <> '000012') " & _
                    "        then (case when (5/cash_avg_day)>1   " & _
                    "             then  1  else (5/cash_avg_day) " & _
                    "             end) " & _
                    "        else 0 end)*pd.mtm) " & _
                    "    /sum(pd.mtm), 5)  " & _
                    "from xrisk.tbsi_ante_port t " & _
                    "left join xrisk.tbsi_ante_portdetail pd on " & _
                    "    pd.pi_id=t.pi_id  " & _
                    "    and pd.w<>0 " & _
                    "    and pd.a_type like '%S%' " & _
                    "where " & _
                    "    t.p_id = '" & GetAccountPID(account) & "' " & _
                    "    and t.bm_id in ('30626', '30724') " & _
                    "    and t.beg_date = '" & Format(infoDate, "yyyy-mm-dd") & "'"
            GetAccountInfo = database.QueryOne(sql, "xRisk")
        Case "止损基准净值"
            If account = "核心组合" Or account = "专项投资" Then
                GetAccountInfo = 30
            Else
                GetAccountInfo = GetAccountInfo(account, infoDate, "净值")
            End If
        Case "止损收益率"
            GetAccountInfo = GetAccountInfo(account, infoDate, "净值") / _
                    GetAccountBaseNV(account) - 1
        Case "止损线"   ' 计算止损线
            If account <> "核心组合" And account <> "专项投资" Then
                GetAccountInfo = 1
            Else
                ' 如果在11月之前，止损线为15%
                If infoDate <= DateValue("2009-11-01") Then
                    GetAccountInfo = 0.15
                Else
                    Dim i As Long, tmpDate As Date, ind As Boolean
                    
                    tmpDate = infoDate
                    ind = False
                    For i = 1 To 10
                        ' 考虑这十个交易日的收益率
                        If GetAccountInfo(account, tmpDate, "止损收益率") < _
                                0.4 - 2 * GetAccountInfo(account, PreTradingday(tmpDate), "止损线") Then
                            GetAccountInfo = GetAccountInfo(account, PreTradingday(tmpDate), "止损线")
                            ind = True
                            Exit For
                        End If
                        
                        tmpDate = PreTradingday(tmpDate)
                    Next i
                    
                    ' 如果这十个交易日的收益率提升收益率，则提升它
                    If ind = False Then
                        GetAccountInfo = -0.05 + GetAccountInfo(account, PreTradingday(infoDate), "止损线")
                    End If
                    
                End If
            End If
    End Select
    
    column = ""
    Select Case LCase(info)
        Case "var"
            column = "var/100000000"
        Case "市值", "mtm", "总市值"
            column = "MTM/100000000"
        Case "成本", "总成本"
            column = "TOTAL_COST/100000000"
        Case "组合收益率", "市值收益率"
            column = "return"
        Case "基准收益率"
            column = "RETURN_BENCH"
        Case "时间加权收益率"
            column = "tw_return"
        Case "alpha", "beta"
            column = "pre_" & info
        Case "beta全市场", "全市场beta"
            column = "pre_beta_bm"
        Case "标准差"
            column = "pre_std"
        Case "基准标准差"
            column = "pre_std_bench"
        Case "事前信息比率"
            column = "pre_inforatio"
    End Select
    
    If Len(column) Then
        sql = "select t." & column & " from xrisk.tbsi_ante_port t " & _
                "where " & _
                "    t.p_id = '" & GetAccountInfo(account, , "pid") & _
                "'    and t.bm_id in ('30626', '30724') " & _
                "    and t.beg_date = '" & Format(infoDate, "yyyy-mm-dd") & "'"
        GetAccountInfo = database.QueryOne(sql, "xRisk")
        
        column = ""
    End If
    
    cache.Add account & infoDate & info, GetAccountInfo
End Function

' 获取某帐户当天的交易情况和持仓情况
' Author: zhangzq@citics, 2009
Public Function GetAccountSecurity(account$, infoDate As Date, _
        Optional securityName$ = "", Optional info$ = "持有市值")
        
    GetAccountSecurity = cache.Item(account & infoDate & securityName & info)
    If GetAccountSecurity <> cache.NotCached Then Exit Function
    
    
    Dim sql$, column$, met$, icode$
    icode = GetSecurityICODE(securityName)
    
    ' 从xRisk获取帐户的交易数据
    met = ""
    Select Case LCase(info)
        Case "交易价格"
            column = "sum(abs(hc_amount))/(0.001+sum(abs(hc_count)))"
        Case "交易金额"
            column = "sum(hc_amount)"
        Case "交易数量"
            column = "sum(hc_count)"
        Case "买入价格"
            column = "trunc(sum(abs(hc_amount))/(0.001+sum(abs(hc_count))), 3)"
            met = " and hc_count > 0 "
        Case "买入金额"
            column = "sum(hc_amount)"
            met = " and hc_count > 0 "
        Case "买入数量"
            column = "sum(hc_count)"
            met = " and hc_count > 0 "
        Case "卖出价格"
            column = "trunc(sum(abs(hc_amount))/(0.001+sum(abs(hc_count))), 3)"
            met = " and hc_count < 0 "
        Case "卖出金额"
            column = "sum(hc_amount)"
            met = " and hc_count < 0 "
        Case "卖出数量"
            column = "sum(hc_count)"
            met = " and hc_count < 0 "
    End Select
    
    If Len(column) Then
        sql = " Select   " & column & _
            " from xrisk.vcrp_chg c " & _
            " where  i_code = '" & icode & "' and " & _
            "    p_id = '" & GetAccountInfo(account, , "pid") & "'" & _
            "    and hc_date='" & Format(infoDate, "yyyy-mm-dd") & "' " & met
        GetAccountSecurity = database.QueryOne(sql, "xRisk")
        column = ""
    End If
    
    ' 从xRisk获取帐户的持仓数据
    Select Case LCase(info)
        Case "持股数量", "存量", "余量"
            column = "sum(h_count)"
        Case "持股成本", "成本", "总成本", "持有成本"
            column = "sum(h_cost)"
        Case "当前市值", "持有市值", "市值"
            column = "sum(h_eval)"
        Case "平均成本", "成本价"
            column = "trunc(sum(h_cost)/(sum(h_count)+0.0000001), 3)"
        Case "市值价格", "收盘价"
            column = "trunc(sum(h_eval)/(sum(h_count)+0.0000001), 3)"
    End Select
    
    If Len(column) Then
        sql = " select " & column & _
            " from  xrisk.vcrp_hld v  " & _
            " where   i_code = '" & icode & "' and " & _
            "    p_id = " & GetAccountInfo(account, , "pid") & _
            "    and v.beg_date ='" & Format(infoDate, "yyyy-mm-dd") & "'"
        GetAccountSecurity = database.QueryOne(sql, "xRisk")
        
        column = ""
    End If
    
    cache.Add account & infoDate & securityName & info, GetAccountSecurity
End Function

' 获取某帐户的某行业当天的持仓情况和交易情况
' Author: zhangzq@citics, 2009
Public Function GetAccountIndustry(account$, infoDate As Date, industry$, Optional info$ = "持有市值")
    GetAccountIndustry = cache.Item(account & infoDate & industry & info)
    If GetAccountIndustry <> cache.NotCached Then Exit Function
        
    Dim columnName$, sql$
    Select Case info
        Case "持股数量", "持有数量"
            columnName = "sum(h_count)/1000000"
        Case "持股成本", "持有成本"
            columnName = "sum(h_cost)/100000000"
        Case "持有市值", "持股市值"
            columnName = "sum(h_eval)/100000000"
    End Select
    
    Select Case industry
        Case "A股", "ETF和LOF基金", "开放式基金", "企业债"
            sql = "select  " & columnName & " " & _
                "from xrisk.vcrp_hld v   " & _
                "left join xrisk.tecd_atype_enum a on  " & _
                "    v.A_TYPE=a.a_type " & _
                "where  " & _
                "    v.p_id = '" & GetAccountInfo(account, infoDate, "pid") & "' " & _
                "    and v.beg_date<='" & Format(infoDate, "yyyy-mm-dd") & "' " & _
                "    and v.end_date>'" & Format(infoDate, "yyyy-mm-dd") & "' " & _
                "    and a.a_name='" & industry & "'"
            GetAccountIndustry = database.QueryOne(sql, "xRisk")
        Case "定向增发"
            GetAccountIndustry = (GetAccountSecurity(account, infoDate, "南  玻Ａ", info) _
                    + GetAccountSecurity(account, infoDate, "金地集团", info) _
                    + GetAccountSecurity(account, infoDate, "中储股份", info)) / 100000000
        Case "1" To "2"
            On Error Resume Next
            GetAccountIndustry = Application.WorksheetFunction.SumIfs(ccSheet.Range(columnName), _
                    ccSheet.Range("B:B"), "=" & account, _
                    ccSheet.Range("A:A"), "=" & Format(infoDate, "yyyy-mm-dd"), _
                    ccSheet.Range("J:J"), "=" & mid(industry, 2))
            If Err.Number > 0 Then Err.Clear: GetAccountIndustry = "#N/A"
        Case Else
            On Error Resume Next
            GetAccountIndustry = Application.WorksheetFunction.SumIfs(ccSheet.Range(columnName), _
                    ccSheet.Range("B:B"), "=" & account, _
                    ccSheet.Range("A:A"), "=" & Format(infoDate, "yyyy-mm-dd"), _
                    ccSheet.Range("I:I"), "=" & industry)
            If Err.Number > 0 Then Err.Clear: GetAccountIndustry = "#N/A"
    End Select
    
    cache.Add account & infoDate & industry & info, GetAccountIndustry
End Function

Private Function GetAccountPID(ByRef account As String) As String
    Select Case account
        Case "交易部整体"
            GetAccountPID = "2628"
        Case "核心组合"
            GetAccountPID = "3201"
        Case "配置组合"
            GetAccountPID = "2629"
        Case "中小盘"
            GetAccountPID = "3867"
        Case "专项投资"
            GetAccountPID = "3844"
        Case "长期战略"
            GetAccountPID = "3453"
        Case Else
            GetAccountPID = "0"
    End Select
End Function

Private Function GetAccountPortCode(account As String) As String
    Select Case account
        Case "核心组合"
            GetAccountPortCode = "79A"
        Case "配置组合"
            GetAccountPortCode = "61A"
        Case "中小盘"
            GetAccountPortCode = "90A"
        Case "专项投资"
            GetAccountPortCode = "88A"
        Case "长期战略"
            GetAccountPortCode = "85A"
        Case Else
            GetAccountPortCode = "0"
    End Select
End Function

Private Function GetAccountBase(ByRef account As String) As String
    Select Case account
        Case "中小盘"
            GetAccountBase = "中小盘复合基准"
        Case Else
            GetAccountBase = "中信标普300"
    End Select
End Function



Public Function GetAccountSetupDate(account As String) As Long
    Select Case account
        Case "交易部整体"
            GetAccountSetupDate = #1/5/2009#
        Case "核心组合"
            GetAccountSetupDate = #1/5/2009#
        Case "配置组合"
            GetAccountSetupDate = #1/5/2009#
        Case "中小盘"
            GetAccountSetupDate = #8/24/2009#
        Case "专项投资"
            GetAccountSetupDate = #8/21/2009#
        Case "长期战略"
            GetAccountSetupDate = #4/16/2009#
        Case Else
            GetAccountSetupDate = 0
    End Select
End Function

Private Function GetAccountBaseNV(account As String) As Double
    Select Case account
        Case "核心组合", "专项投资"
            GetAccountBaseNV = 30.000141
        Case Else
            GetAccountBaseNV = 1
    End Select
End Function


' 获取某个类别的数据
' account: 长期战略
' industry: 某个行业，或者股票，或者整体
' info: 收益率，配置收益率，个股选择收益率等
' beginDate, endDate: 时间区间范围，闭区间
' Author: zhangzq@citics, 2009
Public Function ComputeBF(account As String, industry As String, _
        Optional info As String = "收益率", _
        Optional beginDate As Date, _
        Optional endDate As Date)
    
    If IsMissing(endDate) Then
        endDate = beginDate
    End If
    
    ' 初始化为0
    ComputeBF = 0
    
    ' 计算连乘因子
    Dim kTotal As Double, totalReturn As Double
    totalReturn = GetAccountInfo(account, endDate, "单位净值") / _
            GetAccountInfo(account, PreTradingday(beginDate), "单位净值") - 1
    kTotal = IIf(totalReturn = 0, 1, totalReturn / Log(totalReturn + 1))
    
    Debug.Assert Not kTotal = 0
    
    Dim theDate As Date
    theDate = PreTradingday(endDate + 1)
    Do While theDate >= beginDate
        Dim kDay As Double, dayReturn As Double
        dayReturn = GetAccountInfo(account, theDate, "日收益率")
        kDay = IIf(dayReturn = 0, 1, Log(dayReturn + 1) / dayReturn)
        Debug.Assert Not kDay = 0
        
        ComputeBF = ComputeBF + kDay * kTotal * ComputeDayBF(account, industry, info, theDate)
        
        theDate = PreTradingday(theDate)
    Loop
End Function

' 获得单期数据
' Author: zhangzq@citics, 2009
Public Function ComputeDayBF(account As String, industry As String, _
        Optional info As String = "收益率", _
        Optional infoDate As Date)
    
    Dim column$, sql$
    Select Case info
        Case "组合收益率"
            column = " (case  " & _
                    "        when g.W >0 then g.RETURN   " & _
                    "        else g.RETURN_BENCH " & _
                    "    end)"
        Case "基准收益率"
            column = "g.RETURN_BENCH"
        Case "投资比例"
            column = "g.w"
        Case "基准比例"
            column = "g.w_bench"
        Case "前日比例"
            column = "g.w_pre"
        Case "基准前日比例"
            column = "g.w_pre_bench"
        Case "var"
            column = "g.var"
        Case "beta"
            column = "g.beta"
    End Select
    
    If Len(column) Then
        ' 需要修改
        sql = " SELECT " & column & " " & _
                "from xrisk.tbsi_ante_port p, xrisk.tbsi_ante_group g  " & _
                "left join xrisk.tgrp_def_detail d on " & _
                "    (d.gd_code = g.gd_code " & _
                "    and d.gt_id = g.gt_id) " & _
                "where  " & _
                "    p.p_id = '" & GetAccountInfo(account, , "pid") & "' " & _
                "    and g.pi_id = p.pi_id  " & _
                "    and p.beg_date = '" & Format(infoDate, "yyyy-mm-dd") & "' " & _
                "    and bm_id in ('30626', '30724') " & _
                "    and g.gt_id in ('95')  " & _
                "    and  d.gd_name = '" & industry & "'"
        ComputeDayBF = database.QueryOne(sql, "xRisk")
    End If
    
    
    ' 对于各行业
    Select Case info
        Case "基准配置收益率"
            ComputeDayBF = ComputeDayBF(account, industry, "基准收益率", infoDate) _
                    * ComputeDayBF(account, industry, "基准前日权重", infoDate)
        Case "收益率贡献"
            ComputeDayBF = ComputeDayBF(account, industry, "组合收益率", infoDate) _
                    * ComputeDayBF(account, industry, "前日权重", infoDate)
        Case "超额收益率贡献"
            ComputeDayBF = ComputeDayBF(account, industry, "收益率贡献", infoDate) _
                    - ComputeDayBF(account, industry, "基准配置收益率", infoDate)
        Case "超额配置收益率"
            ComputeDayBF = (ComputeDayBF(account, industry, "基准收益率", infoDate) _
                    - ComputeDayBF(account, "股票", "组合收益率", infoDate)) _
                    * (ComputeDayBF(account, industry, "前日权重", infoDate) _
                    - ComputeDayBF(account, industry, "基准前日权重", infoDate))
        Case "配置收益率"
            ComputeDayBF = ComputeDayBF(account, industry, "基准配置收益率", infoDate) _
                    + ComputeDayBF(account, industry, "超额配置收益率", infoDate)
        Case "个股选择收益率"
            ComputeDayBF = ComputeDayBF(account, industry, "收益率贡献", infoDate) _
                    - ComputeDayBF(account, industry, "配置收益率", infoDate)
    End Select
End Function

' 获取证券的高频价格数据
Public Function GetSecurityPrice(icode$, dt As Date, Optional before As Boolean = True)
    GetSecurityPrice = cache.Item(icode & dt & before)
    If GetSecurityPrice <> cache.NotCached Then Exit Function
    
    Dim sql$
    GetSecurityPrice = database.QueryOne(sql, "gpjy")
    
    cache.Add icode & dt & before, GetSecurityPrice
End Function
