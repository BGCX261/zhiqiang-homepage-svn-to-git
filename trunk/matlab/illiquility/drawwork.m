% sigma=0.01:0.01:2;
% plot(sigma, [illiquityDiscount(1, sigma); ...
%     illiquityDiscount(2, sigma); ...
%     illiquityDiscount(3, sigma)]);

clc
hold off
colormap('Summer');

% 无效数据清零
for i = 1:numel(result)
    if numel(result{i}) == 0
        result{i} = 0;
    end
end

%  大股东
resultR = [result(2:end, :), cell(size(result, 1)-1, 1)];
for i = 2:size(result, 1)
    if strcmp(class(data{i, 11}), 'double')
        resultR{i-1, end} = data{i, 11};
    else
        resultR{i-1, end} = 0;
    end
end



% Longstaff模型的最大折扣率结果
clf
sigma=0.1:0.01:1;
axes('Position', [0.1 0.1 0.85 0.85]);
plot(sigma, illiquityDiscount(1, sigma, 1.02));
hold on
plot(sigma, illiquityDiscount(3, sigma, 1.12), 'color', 'red');
legend('1年期', '3年期');
xlabel('波动率');
ylabel('折扣率上界');
saveas(gcf, 'graph/折扣率上界.jpg');
% 保存为 '折扣率上界.fig'


% 1年期的股票的实证结果
clf
result1 = resultR([resultR{:,4}]<2, :);
sigma=0.01:0.01:1;
plot(sigma, illiquityDiscount(1, sigma, 1.02));
hold on
xlim([0 1]);
ylim([0 1]);
axis([0 1 0 1]);
xlabel('波动率');
ylabel('折扣率');
legend('1年期折扣率上界曲线');
scatter(cell2mat(result1(:, 10)),...
    cell2mat(result1(:, 7)), [], cell2mat(result1(:, 11)), 'filled');
saveas(gcf, 'graph/1年期实际折扣率.jpg');

% 3年期的股票的实证结果
clf
result3 = resultR([resultR{:,4}]>2, :);
sigma=0.01:0.01:1;
plot(sigma, illiquityDiscount(3, sigma, 1.12), 'color', 'red');
hold on
xlim([0 1]);
ylim([0 1]);
axis([0 1 0 1]);
xlabel('波动率');
ylabel('折扣率');
legend('3年期折扣率上界曲线');
scatter(cell2mat(result3(:, 10)),...
    cell2mat(result3(:, 7)), [], cell2mat(result3(:, 11)), 'filled');
saveas(gcf, 'graph/3年期实际折扣率.jpg');


% 增发的股票波动率计算
clf
vol = cell2mat(resultR([resultR{:,10}]>0, 10));
vol = sort(vol);
scatter([1:length(vol)], vol);
hold on
ylim([0.3 1]);
xlim([1 320]);
title('波动率散点图');
ylabel('波动率');
box
saveas(gcf, 'graph/波动率散点图.jpg');

% 上界比例
length(find([result1{:, 7}] < [result1{:, 9}])) ...
    /size(result1, 1)
length(find([result1{:, 7}] < [result1{:, 9}] + 0.1 & ...
    [result1{:, 7}] > [result1{:, 9}] - 0.2)) ...
    /size(result1, 1)


length(find([result3{:, 7}] < [result3{:, 9}])) ...
    /size(result3, 1)

length(find([result1{:, 7}] < [result1{:, 9}] & [result1{:, 11}] < 1)) ...
    /length(find([result1{:, 11}] < 1))
length(find([result3{:, 7}] < [result3{:, 9}] & [result3{:, 11}] < 1)) ...
    /length(find([result3{:, 11}] < 1))
