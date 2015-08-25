function data = databaseQuery(query, type)

setupDatabase

% 默认返回类型为cell
if nargin < 2
   type = 'struct';
end

cur = exec(databaseConn, query);
temp = fetch(cur);
data = temp.Data;

if strcmp(type, 'struct') && strcmpi(query(1:6), 'select')
   databaseName = regexp(query, 'FROM *(\w+)', 'tokens');

   if numel(databaseName)
      databaseName = databaseName{1};
      databaseName = databaseName{1};
      if isfield(databaseColumn, databaseName)
         columnName = getfield(databaseColumn, databaseName);
      else
         columnQuery = ['select   name   from   syscolumns   where  id=object_id(''',databaseName,''') order by colid'];
         columnCur = exec(databaseConn, columnQuery);
         columnTemp = fetch(columnCur);
         columnName = columnTemp.Data;
         databaseColumn = setfield(databaseColumn, databaseName, columnName);
      end
      
      if (ischar(data)) ...
            || numel(data) == 0 || size(data, 2) == 0
         data = [];
      elseif length(columnName) == size(data, 2)
         data = cell2struct(data, columnName, 2);
      else
         cName = {};
         for i = 1:numel(columnName)
            if numel(regexp(query, [columnName{i}, '.*FROM']))
               cName = [cName, columnName(i)];
            end
         end
         if numel(cName) == size(data, 2)
            data = cell2struct(data, cName, 2);
         else
            data = {};
         end
      end
   end
end

