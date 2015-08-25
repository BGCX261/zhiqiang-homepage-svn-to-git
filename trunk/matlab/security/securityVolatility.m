function sigma = securityVolatility(code, dateTo, interval)
% �����Ʊ����ʷ�����ʣ�����ʱ���ΪdateTo֮ǰ��interval��

dateFrom = datestr(datenum(dateTo, 'yyyy-mm-dd') - interval, ...
    'yyyy-mm-dd');
prices = securityWeightedPrice(code, dateFrom, dateTo, 'cell');
% ��ȡ��ʱ��dateFrom��dateTo�ε����й�Ʊ��Ϣ

% ɾ����Ч���ݣ� ĳЩ�۸�ֵΪNaN������˵����ͣ��
% prices = deleteInvalidPrice(prices);
prices = deleteStartInvalidPrice(prices);
prices = deleteEndInvalidPrice(prices);

vol = EstimateVol(interpolatePrice(cell2mat(prices(:,5))),...
    interpolatePrice(cell2mat(prices(:,6))), ...
    interpolatePrice(cell2mat(prices(:,7))), ...
    interpolatePrice(cell2mat(prices(:,8))), size(prices, 1));

sigma = vol.hccv;


function it = interpolatePrice(price) 
% �Լ۸������е���Ч���ݽ��в�ֵ����ֵ������������ʱ����ڽ������ʾ��Ȼ�
% ��Ч����ͨ������Ϊͣ�Ƶ���û�н��׵�����

if strcmp(class(price), 'cell')
    price = cell2mat(price);
end

i = 1;
while ~(price(i) > 0)
    price(i) = [];
end  % ɾ����ʼ������Ч����
while ~(price(end) > 0)
    price(end) = [];
end  % ɾ��������Ч����


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
% ���۸������е���Ч����ɾ��

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
% ���۸������п�ͷ������Ч����ɾ��

ind = false;
i = 1;
while 1
    ind = 1;
    
    for j = 5 : 8
        if strcmp(class(price{i, j}), 'double')  && ((isnan(price{i, j}))...
                || (price{i, j} == 0))
            price(i, :)=[];
            ind = 0;
            % ���������һ��������Ч����ɾ�����У����ı���
            break;
        end
    end
    
    if ind
        break;
    end % �����һ�����ݶ���Ч�� ��ô�������˲�����֤���ݵĵ�һ������Ч�ġ�
end

valid = price;    

function valid = deleteEndInvalidPrice(price)
% ���۸������н�β������Ч����ɾ��

ind = false;
i = 1;
while 1
    ind = 1;
    
    for j = 5 : 8
        if strcmp(class(price{end, j}), 'double')  && ((isnan(price{end, j}))...
                || (price{end, j} == 0))
            price(end, :)=[];
            ind = 0;
            % ���������һ��������Ч����ɾ�����У����ı���
            break;
        end
    end
    
    if ind
        break;
    end % ������һ�����ݶ���Ч�� ��ô�������˲�����֤���ݵ����һ������Ч�ġ�
end

valid = price;    
        









% closePrice = cell2mat(prices(:, 8));
% closePrice = closePrice(closePrice > 0); % ɾ����Ч���ݣ�������ĳЩ�۸�ΪNaN�����ܵ���ͣ��
% 
% rate = log(closePrice);
% rate = rate(2:end) - rate(1:end-1);
% sticks = length(rate);
% 
% if sticks > 1
%     sigma = std(rate)*sqrt(240 - 1); %* sticks / sqrt(sticks - 1);
% else
%     sigma = 0;
%     display(['���������ݲ��������Ʊ', addQuotes(code),'�Ĳ�����', 0]);
% end



    