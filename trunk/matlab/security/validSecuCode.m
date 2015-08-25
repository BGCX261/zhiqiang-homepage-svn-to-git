function validCode = validSecuCode(code)

if strcmp(class(code), 'char')
    validCode = code(1:6);
elseif strcmp(class(code), 'double')
    if code >= 1000000
        error(['�Ƿ�֤ȯ��������: ', num2str(code)]);
    end % �Ƿ�֤ȯ��������
    validCode = sprintf('%06d', code);
end
    