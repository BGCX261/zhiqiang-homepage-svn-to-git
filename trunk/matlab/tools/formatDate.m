function toData = formatDate(fromDate, from, to)
%% ��һ�����ڴ�һ����ʽת������һ����ʽ

toData = datestr(datenum(fromDate, from), to);