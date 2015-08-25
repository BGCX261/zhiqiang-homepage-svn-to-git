function r = illiquityDiscount(T, sigma, rate)
% ʹ��Longstaff�ķ������������ϵĽ����ڲ������ۿ��ʣ� 
% ���룺
%   T ��������
%   sigma ������
%   algorithmType ʹ���㷨��Ĭ��Ϊ1
% ���
%   r �����ۿ���
%
% author & copyright: zhiqiang@citics

if nargin < 3
    rate = -1;
end

temp = sigma.^2 * T ;
r = (2 + temp/2) .* normcdf(sqrt(temp)/2) + sqrt(temp/2/pi) .* exp(-temp/8) - 1;
% Longstaff��ԭʼ�㷨�������ܵ����ۿ��ʴ���100%��

if rate ~=  -1
    r = (rate*(r+1)-1)./(rate*(r+1));
end
% ͨ���ƶ��㷨����algorithmTypeΪ2��ʹ�������������㷨�������ۿ���