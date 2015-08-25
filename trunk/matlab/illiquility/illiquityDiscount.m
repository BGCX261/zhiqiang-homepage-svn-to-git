function r = illiquityDiscount(T, sigma, rate)
% 使用Longstaff的方法计算理论上的禁售期产生的折扣率， 
% 输入：
%   T 禁售年限
%   sigma 波动率
%   algorithmType 使用算法，默认为1
% 输出
%   r 理论折扣率
%
% author & copyright: zhiqiang@citics

if nargin < 3
    rate = -1;
end

temp = sigma.^2 * T ;
r = (2 + temp/2) .* normcdf(sqrt(temp)/2) + sqrt(temp/2/pi) .* exp(-temp/8) - 1;
% Longstaff的原始算法，但可能导致折扣率大于100%！

if rate ~=  -1
    r = (rate*(r+1)-1)./(rate*(r+1));
end
% 通过制定算法类型algorithmType为2，使用我们修正的算法来计算折扣率