global databaseConn;

if ~numel(databaseConn)
    databaseConn = database('jydb','inforead','readinfo');
end

global databaseColumn;
if ~numel(databaseColumn)
    databaseColumn = struct;
end