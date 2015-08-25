clc
clear

addpath('../tools/')
filename = '定向增发-请勿删.xls'; % '定向增发.xls';
% data = xlsread('定向增发.xlsx','固化数据');
data = readXLSToCell('dxzf.xls');
data = data.data;

work = readXLSToCell(filename);
work = work.more1;

now(1, :) = [work(1, :) {'T日'} {'T-1日收盘价'}];

for i = 2 : size(work, 1)
    now(i, :) = [work(i, :), {0}, {0}];
    for j = 2 : size(data, 1)
        if strcmp(data{j, 1}, work{i, 1}) && (data{j, 6}) == (work{i, 14})
            now(i, :) = [work(i, :) data(j, 7:8)];
        end
    end
end



warning off MATLAB:xlswrite:AddSheet
xlswrite(filename, now, 'more2');    % 将结果写入xls文件
