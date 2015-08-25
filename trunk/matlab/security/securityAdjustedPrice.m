function price = securityAdjustedPrice(code, dateFrom, dateTo, type)
% 返回股票code在时间区间dateFrom到dateTo之间(闭区间)的所有交易价格数据
%（按照复权以及大盘指数调整后）。
%
% 输入
%   code 股票代码
%   dateFrom dateTo 时间区间的左端点和右端点 格式 '2001-02-01'，可以为整数，代表天数
%
% author & copyright: zhiqiang@citics

error(nargchk(1, 3, nargin)); % 输入个数为1到3

[innercode, secuMarket] = securityInnercode(code);

if nargin == 2 % 默认只查询一天
    dateTo = 0;
elseif nargin == 1 % 未指定查询日期时查询当天股价
    dateFrom = formatDate(date, 'dd-mmm-yyyy', 'yyyy-mm-dd');
    dateTo = 0;
end

% 获取复权后的价格信息
price = securityWeightedPrice(code, dateFrom, dateTo);

% 获取大盘指数信息
global shIndex;
global szIndex;
if ~numel(shIndex) || ~numel(szIndex)
    shInnercode = securityInnercode('000001', 4);
    szInnercode = securityInnercode('399001', 4);

    shIndex = databaseQuery(['SELECT * FROM QT_DailyQuote ', ...
        'WHERE InnerCode = ', addQuotes(shInnercode), ' order by TradingDay']);

    szIndex = databaseQuery(['SELECT * FROM QT_DailyQuote ', ...
        'WHERE InnerCode = ', addQuotes(szInnercode), ' order by TradingDay']);
end

if secuMarket == 90
    index = szIndex;    % 属于深圳交易所，使用深圳成指
else
    index = shIndex;    % 属于上海交易所，使用上证指数
end

% 针对股市大盘指数进行价格调整
i = 1; j = 1;
while i <= length(price)
    while j < length(index) && strcmpc(price(i).TradingDay, index(j).TradingDay) > 0
        j = j + 1;
    end

    price(i).OpenPrice = price(i).OpenPrice * 1000/index(j).OpenPrice;
    price(i).ClosePrice = price(i).ClosePrice * 1000/index(j).ClosePrice;
    price(i).HighPrice = price(i).HighPrice * 1000/index(j).HighPrice;
    price(i).LowPrice = price(i).LowPrice * 1000/index(j).LowPrice;
    
    i = i + 1;
end

if nargin >= 4 && strcmp(type, 'cell')
    price = squeeze(struct2cell(price))';
end


