function data = queryDatabase(query)

setupDatabase
cur = exec(databaseConn, query);
temp = fetch(cur);
data = temp.Data;

