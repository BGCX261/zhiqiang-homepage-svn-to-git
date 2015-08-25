function price = securityDeleteInvalidPrice(price)

if isstruct(price)
    for i = length(price):-1:1
        if ~isPrice(price(i).ClosePrice) || ~isPrice(price(i).OpenPrice) ...
               || ~isPrice(price(i).LowPrice) || ~isPrice(price(i).HighPrice) ...
               || ~isPrice(price(i).TurnoverValue) || ~isPrice(price(i).TurnoverVolume)
            price(i) = [];
        end
    end
end

function r = isPrice(p)

if isnumeric(p) && p > 0
    r = true;
else
    r = false;
end
