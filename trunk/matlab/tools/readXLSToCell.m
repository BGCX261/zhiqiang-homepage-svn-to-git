function [data, status1, status2] = readXLSToCell(filename, sheet)

[file, status1] = officedoc(filename, 'open', 'mode','read');
data = officedoc(file, 'read');
status2 = officedoc(file, 'close');