function price = securityWeightedPrice(code, dateFrom, dateTo, type, field)
% 返回股票code在时间区间dateFrom到dateTo之间(闭区间)的所有交易价格数据（按照复权调整后）。
%
% 输入
%   code 股票代码
%   dateFrom dateTo 时间区间的左端点和右端点 格式 '2001-02-01'，可以为整数，代表天数
%
% author & copyright: zhiqiang@citics


error(nargchk(1, 5, nargin)); % 输入个数为1到4
if nargin < 5
   field = ' * ';
end

innercode = securityInnercode(code);

if nargin == 2 % 默认只查询一天
   dateTo = 0;
elseif nargin == 1 % 未指定查询日期时查询当天股价
   dateFrom = formatDate(date, 'dd-mmm-yyyy', 'yyyy-mm-dd');
   dateTo = 0;
end

% 获取实际价格信息
price = securityPrice(code, dateFrom, dateTo, 'struct', field);
% 获取复权信息
weight = databaseQuery(['SELECT * FROM QT_AdjustingFactor ', ...
   'WHERE InnerCode = ', addQuotes(innercode),...
   ' AND ExdiviDate <= ', addQuotes(dateTo)]);

if ~isstruct(weight) || numel(weight) == 1

else % 如果数据库有复权信息， 对价格进行复权操作
   % 以下这一段计算复权价格
   weightNow = 1;
   i = 1; j = 1;
   while i <= length(price);
      % 查找合适的复权， 它是不大于当前时间且最近的复权因子

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

      % 对此复权，应用到所有小于下一个复权时间的所有价格
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