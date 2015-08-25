clc
clear

addpath('../tools/')
filename = 'all.xls'; % '��������.xls';
% data = xlsread('��������.xlsx','�̻�����');
data = readXLSToCell(filename);
data = data.data;

result = struct;

result(1).code='����';
result(1).name='����';
result(1).lockup = '������';
result(1).announceDate='��׼��';
result(1).announceDatePrice='��׼�ռ۸�';
result(1).furtherDate='������';
result(1).actualPrice='�����ռ۸�';
result(1).furtherPrice='�����۸�';
result(1).dateInterval = '���ʱ��';
result(1).actualToAnnouce='��Ʊ�����ʣ�δ�����г����أ�';
result(1).adjustedActualToAnnounceDate='��Ʊ������(����ڴ���)';
result(1).announceDateDiscount='��׼���ۿ���';
result(1).accountDateDiscount = '�������ۿ���';
result(1).furtherDateDiscount = 'ʵ���ۿ���';
result(1).theoryDiscount = '�����ۿ���';
result(1).volatility = '������';

for i = 2 : size(data, 1)
    secu = data(i, :);
    result(i).code = secu{1};	% ��Ʊ��˾����
    result(i).name = secu{2};	% ��Ʊ��˾����
    result(i).lockup = (datenum(secu{18}, 'yyyymmdd') ...
        - datenum(secu{17}, 'yyyymmdd'))/365; % ������

    furtherDate = formatDate(secu{17}, 'yyyymmdd', 'yyyy-mm-dd');
    result(i).furtherDate = furtherDate;     % �����ɷ���������
    furtherPrice = str2double(secu{13});
    result(i).furtherPrice = furtherPrice;    % �����۸�

    newIssue = AShareSeasonedNewIssue(secu{1}, ['IssuePrice = ', secu{13}]);
    if numel(newIssue) ~= 1
        continue;
    end


    result(i).announceDate = newIssue.InitialInfoPublDate;    % ���»ṫ��������

    result(i).dateInterval = (datenum(result(i).furtherDate, 'yyyy-mm-dd') ...
        - datenum(result(i).announceDate, 'yyyy-mm-dd'))/365;  % �������
    result(i).announceDatePrice = ...
        result(i).furtherPrice * 100 / secu{21};   % �����»ṫ���գ���׼�۸�

    % queryBeginDate = datestr(datenum(furtherDate, 'yyyy-mm-dd'), 'yyyy-mm-dd');
    actualPrice = securityPrice(secu{1}, -1, furtherDate);
    if numel(actualPrice) ~= 1
        result(i).actualPrice = actualPrice{1, 8};    % ��������ʵ�ʼ۸�
    end

    result(i).announceDateDiscount = secu{21};        % ��׼���ۿ���
    result(i).furtherDateDiscount = 1 - result(i).furtherPrice ...
        /result(i).actualPrice; % ʵ���ۿ���

    accountDate = newIssue.DateToAccount;
    accountDatePrice = securityPrice(secu{1}, -20, accountDate, 'struct');
    if numel(accountDatePrice)
        accountPrice = sum([accountDatePrice.TurnoverValue]) ...
            /sum([accountDatePrice.TurnoverVolume]);
        result(i).accountDateDiscount = 1 - result(i).furtherPrice ...
            /accountPrice; % ���ڵ����չɼ۵��ۿ���
    end
    
    % ��Ȩ�ɼ������ʣ������գ�������
    announceDateWeightedPrice = securityWeightedPrice(secu{1},...
        -1, result(i).announceDate, 'struct');
    actualWeightedPrice = securityWeightedPrice(secu{1}, ...
        1, furtherDate, 'struct');
    if numel(announceDateWeightedPrice) && numel(actualWeightedPrice)
        result(i).actualToAnnouce = actualWeightedPrice.ClosePrice / ...
            announceDateWeightedPrice.ClosePrice * 100 - 100;
    end

    % ��Ȩ�ɼ�����ڴ��̵������ʣ������գ�������
    adjustedAnnounceDatePrice = securityAdjustedPrice(secu{1}, ...
        -1, result(i).announceDate);
    adjustedActualPrice = securityAdjustedPrice(secu{1}, 1, furtherDate);
    if numel(adjustedAnnounceDatePrice) && numel(adjustedActualPrice)
        result(i).adjustedActualToAnnounceDate = ...
            adjustedActualPrice(1).ClosePrice / ...
            adjustedAnnounceDatePrice(1).ClosePrice * 100 - 100;
    end

    result(i).volatility = securitySigma(secu{1}, furtherDate, ...
        datenum(secu{18}, 'yyyymmdd') - ...
        datenum(secu{17}, 'yyyymmdd'));  % ��Ʊ�Ĳ�����

    rate = RMBInterestRate(furtherDate);
    if result(i).lockup < 2
        rate = exp(result(i).lockup*(rate{8}-rate{5}));
    else
        rate = exp(result(i).lockup*(rate{10}-rate{5}));
    end
    result(i).theoryDiscount = illiquityDiscount(result(i).lockup, ...
        result(i).volatility, rate);         % ������������ۿ���
    %   result{i, 8} = illiquityDiscount(result{i, 4}, result{i, 10});  % �����ۿ���


end

warning off MATLAB:xlswrite:AddSheet
xlswrite(filename, squeeze(struct2cell(result))', 'datalimit');    % �����д��xls�ļ�









% % ���������Թɼ۵�Ӱ��
% pr = zeros(size(data, 1) - 1, 3);
% for i = 2 : size(data, 1)
%     secu = data(i, :);
%     furtherDate = formatDate(secu{17}, 'yyyymmdd', 'yyyy-mm-dd');
%     tempPrice = [];
%     tempPrice = securityPrice(secu{1}, furtherDate, -30);
%     if strcmp(class(tempPrice), 'cell')
%         pr(i-1, 1) = mean(cell2mat(tempPrice(:, 8)));
%     end
%     postDate = formatDate(secu{4}, 'yyyymmdd', 'yyyy-mm-dd');
%     tempPrice = [];
%     tempPrice = securityPrice(secu{1}, postDate, -30);
%     if strcmp(class(tempPrice), 'cell')
%         pr(i-1, 2) = mean(cell2mat(tempPrice(:, 8)));
%     end
% end
%



