clc
clear

addpath('../tools/')
filename = '��������-����ɾ.xls'; % '��������.xls';
% data = xlsread('��������.xlsx','�̻�����');
data = readXLSToCell('dxzf.xls');
data = data.data;

work = readXLSToCell(filename);
work = work.more1;

now(1, :) = [work(1, :) {'T��'} {'T-1�����̼�'}];

for i = 2 : size(work, 1)
    now(i, :) = [work(i, :), {0}, {0}];
    for j = 2 : size(data, 1)
        if strcmp(data{j, 1}, work{i, 1}) && (data{j, 6}) == (work{i, 14})
            now(i, :) = [work(i, :) data(j, 7:8)];
        end
    end
end



warning off MATLAB:xlswrite:AddSheet
xlswrite(filename, now, 'more2');    % �����д��xls�ļ�
