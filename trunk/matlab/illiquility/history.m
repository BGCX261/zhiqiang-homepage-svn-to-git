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
result(1).announceDate='基准日';
result(1).announceDatePrice='基准日价格';
result(1).furtherDate='上市日';
result(1).actualPrice='上市日价格';
result(1).furtherPrice='增发价格';
result(1).dateInterval = '间隔时间';
result(1).announceDateAccount='基准日折扣率';
result(1).actualToAnnouce='股票增长率（未考虑市场因素）';
result(1).adjustedActualToAnnounceDate='股票增长率(相对于大盘)';

for i = 2 : size(data, 1)
    secu = data(i, :);
    result(i).code = secu{1};	% 股票公司代码
    result(i).name = secu{2};	% 股票公司名字

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
    actualPrice = securityPrice(secu{1}, furtherDate);
    if numel(actualPrice) ~= 1
        result(i).actualPrice = actualPrice{1, 8};    % 增发当天实际价格
    end

    result(i).announceDateAccount = secu{21};        % 基准日折扣率
    result(i).actualToAnnouce = result(i).actualPrice / ...
        result(i).announceDatePrice * 100;   % 股价上涨率

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
end

warning off MATLAB:xlswrite:AddSheet
xlswrite(filename, squeeze(struct2cell(result))', 'history');    % 将结果写入xls文件