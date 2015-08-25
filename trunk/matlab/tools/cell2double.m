function d = cell2double(c, type)
% 将元胞c里的所有double型数据转化成一个实数矩阵。

if nargin < 2
    type = 2
end

d = [];

for i = 1:length(c(:))
    if strcmp(class(c{i}), 'double')
        d = [d c{i}];
    end
end