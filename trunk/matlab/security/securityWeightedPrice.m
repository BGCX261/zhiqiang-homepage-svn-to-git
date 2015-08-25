function price = securityWeightedPrice(code, dateFrom, dateTo, type, field)
% ���ع�Ʊcode��ʱ������dateFrom��dateTo֮��(������)�����н��׼۸����ݣ����ո�Ȩ�����󣩡�
%
% ����
%   code ��Ʊ����
%   dateFrom dateTo ʱ���������˵���Ҷ˵� ��ʽ '2001-02-01'������Ϊ��������������
%
% author & copyright: zhiqiang@citics


error(nargchk(1, 5, nargin)); % �������Ϊ1��4
if nargin < 5
   field = ' * ';
end

innercode = securityInnercode(code);

if nargin == 2 % Ĭ��ֻ��ѯһ��
   dateTo = 0;
elseif nargin == 1 % δָ����ѯ����ʱ��ѯ����ɼ�
   dateFrom = formatDate(date, 'dd-mmm-yyyy', 'yyyy-mm-dd');
   dateTo = 0;
end

% ��ȡʵ�ʼ۸���Ϣ
price = securityPrice(code, dateFrom, dateTo, 'struct', field);
% ��ȡ��Ȩ��Ϣ
weight = databaseQuery(['SELECT * FROM QT_AdjustingFactor ', ...
   'WHERE InnerCode = ', addQuotes(innercode),...
   ' AND ExdiviDate <= ', addQuotes(dateTo)]);

if ~isstruct(weight) || numel(weight) == 1

else % ������ݿ��и�Ȩ��Ϣ�� �Լ۸���и�Ȩ����
   % ������һ�μ��㸴Ȩ�۸�
   weightNow = 1;
   i = 1; j = 1;
   while i <= length(price);
      % ���Һ��ʵĸ�Ȩ�� ���ǲ����ڵ�ǰʱ��������ĸ�Ȩ����

      try
         j <= length(weight) && strcmpc(price(i).TradingDay, ...
            weight(j).ExDiviDate) >= 0;
      catch
         class(weight), size(weight)
      end

      while j <= length(weight) && strcmpc(price(i).TradingDay, ...
            weight(j).ExDiviDate) >= 0
         weightNow = weight(j).AdjustingFactor;
         j = j + 1;
      end

      % �Դ˸�Ȩ��Ӧ�õ�����С����һ����Ȩʱ������м۸�
      while i <= length(price) && ...
            (j > length(weight) || strcmpc(price(i).TradingDay, ...
            weight(j).ExDiviDate) < 0)
         if isfield(price(1), 'OpenPrice')
            price(i).OpenPrice = price(i).OpenPrice * weightNow;
         end
         if isfield(price(1), 'ClosePrice')
            price(i).ClosePrice = price(i).ClosePrice * weightNow;
         end
         if isfield(price(1), 'HighPrice')
            price(i).HighPrice = price(i).HighPrice * weightNow;
         end
         if isfield(price(1), 'LowPrice')
            price(i).LowPrice = price(i).LowPrice * weightNow;
         end
         i = i + 1;
      end
   end
end

if nargin >= 4 && strcmp(type, 'cell')
   price = squeeze(struct2cell(price))';
end