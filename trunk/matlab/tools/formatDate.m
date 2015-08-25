function toData = formatDate(fromDate, from, to)
%% 将一个日期从一种形式转化到另一个形式

toData = datestr(datenum(fromDate, from), to);