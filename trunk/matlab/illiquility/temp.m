clc
clear

addpath('../tools/')
filename = 'all.xls'; % '定向增发.xls';
% data = xlsread('定向增发.xlsx','固化数据');
data = readXLSToCell(filename);
data = data.data;

result = struct;

result(1).code='代码';
result(1).name='名称';
result(1).lockup = '禁售期';
result(1).announceDate='基准日';
result(1).announceDatePrice='基准日价格';
result(1).furtherDate='上市日';
result(1).actualPrice='上市日价格';
result(1).furtherPrice='增发价格';
result(1).dateInterval = '间隔时间';
result(1).actualToAnnouce='股票增长率（未考虑市场因素）';
result(1).adjustedActualToAnnounceDate='股票增长率(相对于大盘)';
result(1).announceDateDiscount='基准日折扣率';
result(1).accountDateDiscount = '到帐日折扣率';
result(1).furtherDateDiscount = '实际折扣率';
result(1).theoryDiscount = '理论折扣率';
result(1).volatility = '波动率';

for i = 2 : size(data, 1)
    secu = data(i, :);
    result(i).code = secu{1};	% 股票公司代码
    result(i).name = secu{2};	% 股票公司名字
    result(i).lockup = (datenum(secu{18}, 'yyyymmdd') ...
        - datenum(secu{17}, 'yyyymmdd'))/365; % 禁售期

    furtherDate = formatDate(secu{17}, 'yyyymmdd', 'yyyy-mm-dd');
    result(i).furtherDate = furtherDate;     % 增发股份上市日期
    furtherPrice = str2double(secu{13});
    result(i).furtherPrice = furtherPrice;    % 增发价格

    newIssue = AShareSeasonedNewIssue(secu{1}, ['IssuePrice = ', secu{13}]);
    if numel(newIssue) ~= 1
        continue;
    end


    result(i).announceDate = newIssue.InitialInfoPublDate;    % 董事会公告日日期

    result(i).dateInterval = (datenum(result(i).furtherDate, 'yyyy-mm-dd') ...
        - datenum(result(i).announceDate, 'yyyy-mm-dd'))/365;  % 间隔日期
    result(i).announceDatePrice = ...
        result(i).furtherPrice * 100 / secu{21};   % （董事会公告日）基准价格

    % queryBeginDate = datestr(datenum(furtherDate, 'yyyy-mm-dd'), 'yyyy-mm-dd');
    actualPrice = securityPrice(secu{1}, -1, furtherDate);
    if numel(actualPrice) ~= 1
        result(i).actualPrice = actualPrice{1, 8};    % 增发当天实际价格
    end

    result(i).announceDateDiscount = secu{21};        % 基准日折扣率
    result(i).furtherDateDiscount = 1 - result(i).furtherPrice ...
        /result(i).actualPrice; % 实际折扣率

    accountDate = newIssue.DateToAccount;
    accountDatePrice = securityPrice(secu{1}, -20, accountDate, 'struct');
    if numel(accountDatePrice)
        accountPrice = sum([accountDatePrice.TurnoverValue]) ...
            /sum([accountDatePrice.TurnoverVolume]);
        result(i).accountDateDiscount = 1 - result(i).furtherPrice ...
            /accountPrice; % 基于到帐日股价的折扣率
    end
    
    % 复权股价增长率，增发日：公告日
    announceDateWeightedPrice = securityWeightedPrice(secu{1},...
        -1, result(i).announceDate, 'struct');
    actualWeightedPrice = securityWeightedPrice(secu{1}, ...
        1, furtherDate, 'struct');
    if numel(announceDateWeightedPrice) && numel(actualWeightedPrice)
        result(i).actualToAnnouce = actualWeightedPrice.ClosePrice / ...
            announceDateWeightedPrice.ClosePrice * 100 - 100;
    end

    % 复权股价相对于大盘的增长率，增发日：公告日
    adjustedAnnounceDatePrice = securityAdjustedPrice(secu{1}, ...
        -1, result(i).announceDate);
    adjustedActualPrice = securityAdjustedPrice(secu{1}, 1, furtherDate);
    if numel(adjustedAnnounceDatePrice) && numel(adjustedActualPrice)
        result(i).adjustedActualToAnnounceDate = ...
            adjustedActualPrice(1).ClosePrice / ...
            adjustedAnnounceDatePrice(1).ClosePrice * 100 - 100;
    end

    result(i).volatility = securitySigma(secu{1}, furtherDate, ...
        datenum(secu{18}, 'yyyymmdd') - ...
        datenum(secu{17}, 'yyyymmdd'));  % 股票的波动率

    rate = RMBInterestRate(furtherDate);
    if result(i).lockup < 2
        rate = exp(result(i).lockup*(rate{8}-rate{5}));
    else
        rate = exp(result(i).lockup*(rate{10}-rate{5}));
    end
    result(i).theoryDiscount = illiquityDiscount(result(i).lockup, ...
        result(i).volatility, rate);         % 修正后的理论折扣率
    %   result{i, 8} = illiquityDiscount(result{i, 4}, result{i, 10});  % 理论折扣率


end

warning off MATLAB:xlswrite:AddSheet
xlswrite(filename, squeeze(struct2cell(result))', 'datalimit');    % 将结果写入xls文件









% % 定向增发对股价的影响
% pr = zeros(size(data, 1) - 1, 3);
% for i = 2 : size(data, 1)
%     secu = data(i, :);
%     furtherDate = formatDate(secu{17}, 'yyyymmdd', 'yyyy-mm-dd');
%     tempPrice = [];
%     tempPrice = securityPrice(secu{1}, furtherDate, -30);
%     if strcmp(class(tempPrice), 'cell')
%         pr(i-1, 1) = mean(cell2mat(tempPrice(:, 8)));
%     end
%     postDate = formatDate(secu{4}, 'yyyymmdd', 'yyyy-mm-dd');
%     tempPrice = [];
%     tempPrice = securityPrice(secu{1}, postDate, -30);
%     if strcmp(class(tempPrice), 'cell')
%         pr(i-1, 2) = mean(cell2mat(tempPrice(:, 8)));
%     end
% end
%



