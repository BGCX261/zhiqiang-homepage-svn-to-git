function price = securityAdjustedPrice(code, dateFrom, dateTo, type)
% ���ع�Ʊcode��ʱ������dateFrom��dateTo֮��(������)�����н��׼۸�����
%�����ո�Ȩ�Լ�����ָ�������󣩡�
%
% ����
%   code ��Ʊ����
%   dateFrom dateTo ʱ���������˵���Ҷ˵� ��ʽ '2001-02-01'������Ϊ��������������
%
% author & copyright: zhiqiang@citics

error(nargchk(1, 3, nargin)); % �������Ϊ1��3

[innercode, secuMarket] = securityInnercode(code);

if nargin == 2 % Ĭ��ֻ��ѯһ��
    dateTo = 0;
elseif nargin == 1 % δָ����ѯ����ʱ��ѯ����ɼ�
    dateFrom = formatDate(date, 'dd-mmm-yyyy', 'yyyy-mm-dd');
    dateTo = 0;
end

% ��ȡ��Ȩ��ļ۸���Ϣ
price = securityWeightedPrice(code, dateFrom, dateTo);

% ��ȡ����ָ����Ϣ
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
    index = szIndex;    % �������ڽ�������ʹ�����ڳ�ָ
else
    index = shIndex;    % �����Ϻ���������ʹ����ָ֤��
end

% ��Թ��д���ָ�����м۸����
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


