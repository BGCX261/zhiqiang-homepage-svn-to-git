function [data, status1, status2] = readXLSToCell(filename, sheet)
% ����һ��xls�ļ���cell
% 
% data = readXLSToCell(filename) ��ȡfilename(ֻ֧��xls�ļ���
% ��֧��office 2007��xlsx��ʽ)������һ���ṹ���ṹ��ÿ����
% Ա����һ��cell���󣬴���XLS�ļ����һ�ű�
%
% data = readXLSToCell(filename, num) ��ȡfilename�ĵ�num�ű�
% data�ڣ�data����һ��cell����
% 
% data = readXLSToCell(filename, sheetname) ��ȡfilename��sheetname
% ��data�ڣ�data����һ��cell����
%  
% ע��filename��sheetname��֧�����ġ�
% ��ע��matlab�Դ���xlswrite(filename, data, sheetname)��
% �԰�cell����dataд�뵽�ļ�filename�ı�sheetname�
% ��������������Ͽɻ�������xls��matlab֮������ݴ�������
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
