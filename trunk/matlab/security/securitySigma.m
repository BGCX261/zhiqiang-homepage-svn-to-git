function sigma = securityVolatility(code, dateTo, interval)
% 计算股票的历史波动率，计算时间段为dateTo之前的interval天

dateFrom = datestr(datenum(dateTo, 'yyyy-mm-dd') - interval, ...
    'yyyy-mm-dd');
prices = securityWeightedPrice(code, dateFrom, dateTo, 'cell');
% 获取从时间dateFrom到dateTo段的所有股票信息

% 删除无效数据， 某些价格值为NaN，比如说当天停牌
% prices = deleteInvalidPrice(prices);
prices = deleteStartInvalidPrice(prices);
prices = deleteEndInvalidPrice(prices);

vol = EstimateVol(interpolatePrice(cell2mat(prices(:,5))),...
    interpolatePrice(cell2mat(prices(:,6))), ...
    interpolatePrice(cell2mat(prices(:,7))), ...
    interpolatePrice(cell2mat(prices(:,8))), size(prices, 1));

sigma = vol.hccv;


function it = interpolatePrice(price) 
% 对价格数据中的无效数据进行插值。插值方法是在连续时间段内将收益率均匀化
% 无效数据通常是因为停牌当天没有交易等因素

if strcmp(class(price), 'cell')
    price = cell2mat(price);
end

i = 1;
while ~(price(i) > 0)
    price(i) = [];
end  % 删除开始处的无效数据
while ~(price(end) > 0)
    price(end) = [];
end  % 删除最后的无效数据


while i < length(price)
    if price(i) > 0
        i = i + 1;
    else
        j = i + 1;
        while j < length(price) && ( ~(price(j) > 0) )
            j = j + 1;
        end
       
        p = (price(j)/price(i-1))^(1/(j-i+1));
        
        for k = i : j-1
            price(k) = price(k-1) * p;
        end
        
        i = j + 1;
    end
end

it = price;



function valid = deleteInvalidPrice(price)
% 将价格数据中的无效数据删除

for i = size(price, 1) : -1 : 1
    for j = 5 : 8
        if strcmp(class(price{i, j}), 'double')  && ((isnan(price{i, j})) || (price{i, j} == 0))
            price(i, :)=[];
            break;
        end
    end
end

valid = price;    

function valid = deleteStartInvalidPrice(price)
% 将价格数据中开头处的无效数据删除

ind = false;
i = 1;
while 1
    ind = 1;
    
    for j = 5 : 8
        if strcmp(class(price{i, j}), 'double')  && ((isnan(price{i, j}))...
                || (price{i, j} == 0))
            price(i, :)=[];
            ind = 0;
            % 如果此行有一个数据无效，则删除整行，并改变标记
            break;
        end
    end
    
    if ind
        break;
    end % 如果第一行数据都有效， 那么跳出。此操作保证数据的第一行是有效的。
end

valid = price;    

function valid = deleteEndInvalidPrice(price)
% 将价格数据中结尾处的无效数据删除

ind = false;
i = 1;
while 1
    ind = 1;
    
    for j = 5 : 8
        if strcmp(class(price{end, j}), 'double')  && ((isnan(price{end, j}))...
                || (price{end, j} == 0))
            price(end, :)=[];
            ind = 0;
            % 如果此行有一个数据无效，则删除整行，并改变标记
            break;
        end
    end
    
    if ind
        break;
    end % 如果最后一行数据都有效， 那么跳出。此操作保证数据的最后一行是有效的。
end

valid = price;    
        









% closePrice = cell2mat(prices(:, 8));
% closePrice = closePrice(closePrice > 0); % 删除无效数据，数据中某些价格为NaN，可能当天停牌
% 
% rate = log(closePrice);
% rate = rate(2:end) - rate(1:end-1);
% sticks = length(rate);
% 
% if sticks > 1
%     sigma = std(rate)*sqrt(240 - 1); %* sticks / sqrt(sticks - 1);
% else
%     sigma = 0;
%     display(['区间内数据不够计算股票', addQuotes(code),'的波动率', 0]);
% end



    