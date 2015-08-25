function newIssue = AShareSeasonedNewIssue(code, condition)

innercode = securityInnercode(code);
newIssue = databaseQuery(['SELECT * FROM LC_AShareSeasonedNewIssue ', ...
    'WHERE InnerCode = ', addQuotes(innercode), ' AND ', condition]);