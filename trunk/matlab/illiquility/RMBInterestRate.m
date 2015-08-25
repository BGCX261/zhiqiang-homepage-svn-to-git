function rate = RMBInterestRate(queryDay)

rate = queryDatabase(['SELECT TOP 1 * FROM Bond_RMBInterestRate ', ...
    'WHERE AdjustDate <= ', ...
    addQuotes(queryDay), ' ORDER BY ID DESC']);

% queryDatabase(['SELECT COL_NAME FROM Bond_RMBInterestRate']); 