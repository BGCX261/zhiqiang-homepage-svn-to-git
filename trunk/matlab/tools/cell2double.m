function d = cell2double(c, type)
% ��Ԫ��c�������double������ת����һ��ʵ������

if nargin < 2
    type = 2
end

d = [];

for i = 1:length(c(:))
    if strcmp(class(c{i}), 'double')
        d = [d c{i}];
    end
end