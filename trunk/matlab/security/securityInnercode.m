function [innercode, secuMarket] = securityInnercode(ocode, secuCategory)

code = validSecuCode(ocode);

if nargin == 1
    secuCategory = 1;
end

if secuCategory < 0 || ...
      (ischar(ocode) && strcmpc(ocode(end-2:end), '.ZS') == 0)
   categoryQuery = [' AND SecuCategory = 4'];
else
   categoryQuery = [' AND SecuCategory = ', addQuotes(secuCategory)];
end

temp = queryDatabase(['SELECT InnerCode, SecuMarket FROM SecuMain ', ...
    'WHERE SecuCode = ', addQuotes(code), categoryQuery]);
if length(temp) == 2
    innercode = temp{1};
    secuMarket = temp{2};
else
    innercode = temp;
end