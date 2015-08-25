function [data, status1, status2] = readXLSToCell(filename, sheet)
% 读入一个xls文件到cell
% 
% data = readXLSToCell(filename) 读取filename(只支持xls文件，
% 不支持office 2007的xlsx格式)，返回一个结构，结构的每个成
% 员都是一个cell矩阵，代表XLS文件里的一张表。
%
% data = readXLSToCell(filename, num) 读取filename的第num张表到
% data内，data将是一个cell矩阵。
% 
% data = readXLSToCell(filename, sheetname) 读取filename的sheetname
% 表到data内，data将是一个cell矩阵。
%  
% 注：filename和sheetname都支持中文。
% 另注：matlab自带的xlswrite(filename, data, sheetname)可
% 以把cell矩阵data写入到文件filename的表sheetname里，
% 这两个函数相配合可基本满足xls与matlab之间的数据传递需求。
%
% copyright: zhang@zhiqiang.org, last update: 2009-08-07

[file, status1] = officedoc(filename, 'open', 'mode','read');
data = officedoc(file, 'read');
status2 = officedoc(file, 'close');

if nargin == 2 && ischar(sheet)
    data = data.(chinese2english(sheet));
elseif nargin == 2 && isnumeric(sheet)
    fields = fieldnames(data);
    if sheet > numel(fields)
        error('not enough sheets');
    end
    fields = fields{sheet};
    data = data.(fields);
end
