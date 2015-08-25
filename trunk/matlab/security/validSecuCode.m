function validCode = validSecuCode(code)

if strcmp(class(code), 'char')
    validCode = code(1:6);
elseif strcmp(class(code), 'double')
    if code >= 1000000
        error(['非法证券代码输入: ', num2str(code)]);
    end % 非法证券代码输入
    validCode = sprintf('%06d', code);
end
    