function price = securityPrice(code, dateFrom, dateTo, type, field)
% ���ع�Ʊcode��ʱ������dateFrom��dateTo֮��(������)�����н��׼۸����ݡ�
%
% ����
%   code ��Ʊ����
%   dateFrom dateTo ʱ���������˵���Ҷ˵� ��ʽ '2001-02-01'������Ϊ��������������
%
% author & copyright: zhiqiang@citics

error(nargchk(1, 5, nargin)); % �������Ϊ1��4

innercode = securityInnercode(code);

if nargin == 2 % Ĭ��ֻ��ѯһ��
    dateTo = 0;
elseif nargin == 1 % δָ����ѯ����ʱ��ѯ����ɼ�
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
end  % ����������͵����룬��߿�����

if nargin < 4 || ~strcmp(type, 'struct')
    price = queryDatabase(['SELECT ', field, ' FROM QT_DailyQuote ', ...
        'WHERE InnerCode = ', addQuotes(innercode), 'AND  TurnoverVolume > 0 AND TradingDay >= ', ...
        addQuotes(dateFrom), ' AND TradingDay <= ', addQuotes(dateTo)]);
elseif strcmp(type, 'struct')
    price = databaseQuery(['SELECT ', field, ' FROM QT_DailyQuote ', ...
        'WHERE InnerCode = ', addQuotes(innercode), 'AND  TurnoverVolume > 0 AND TradingDay >= ', ...
        addQuotes(dateFrom), ' AND TradingDay <= ', addQuotes(dateTo)]);
end