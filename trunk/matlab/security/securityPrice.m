function price = securityPrice(code, dateFrom, dateTo, type, field)
% 返回股票code在时间区间dateFrom到dateTo之间(闭区间)的所有交易价格数据。
%
% 输入
%   code 股票代码
%   dateFrom dateTo 时间区间的左端点和右端点 格式 '2001-02-01'，可以为整数，代表天数
%
% author & copyright: zhiqiang@citics

error(nargchk(1, 5, nargin)); % 输入个数为1到4

innercode = securityInnercode(code);

if nargin == 2 % 默认只查询一天
    dateTo = 0;
elseif nargin == 1 % 未指定查询日期时查询当天股价
    dateFrom = formatDate(date, 'dd-mmm-yyyy', 'yyyy-mm-dd');
    dateTo = 0;
end

if nargin < 5
   field = ' * ';
elseif field == -1
   field = ' TradingDay, ClosePrice ';
else
   field = [' ' field ' '];
end
  

if strcmp(class(dateTo), 'double')
    dateTo = datestr(datenum(dateFrom, 'yyyy-mm-dd') + dateTo, 'yyyy-mm-dd');
    if strcmpc(dateFrom, dateTo) > 0
        dateT = dateTo;
        dateTo = dateFrom;
        dateFrom = dateT;
    end
elseif strcmp(class(dateFrom), 'double')
    if dateFrom > 0
        query = ['SELECT TOP ', num2str(dateFrom), field,  ' FROM QT_DailyQuote ', ...
            'WHERE TurnoverVolume > 0 AND InnerCode = ', addQuotes(innercode), 'AND TradingDay >= ', ...
            addQuotes(dateTo)];
    else
        query = ['SELECT TOP ', num2str(abs(dateFrom)), field,  ' FROM QT_DailyQuote ', ...
            'WHERE  TurnoverVolume > 0 AND InnerCode = ', addQuotes(innercode), 'AND TradingDay <= ', ...
            addQuotes(dateTo), ' ORDER BY TradingDay DESC'];
    end
    if nargin < 4 || ~strcmp(type, 'struct')
        price = queryDatabase(query);
    elseif strcmp(type, 'struct')
        price = databaseQuery(query);
    end
    
    return;
end  % 处理各种类型的输入，提高可用性

if nargin < 4 || ~strcmp(type, 'struct')
    price = queryDatabase(['SELECT ', field, ' FROM QT_DailyQuote ', ...
        'WHERE InnerCode = ', addQuotes(innercode), 'AND  TurnoverVolume > 0 AND TradingDay >= ', ...
        addQuotes(dateFrom), ' AND TradingDay <= ', addQuotes(dateTo)]);
elseif strcmp(type, 'struct')
    price = databaseQuery(['SELECT ', field, ' FROM QT_DailyQuote ', ...
        'WHERE InnerCode = ', addQuotes(innercode), 'AND  TurnoverVolume > 0 AND TradingDay >= ', ...
        addQuotes(dateFrom), ' AND TradingDay <= ', addQuotes(dateTo)]);
end