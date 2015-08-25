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
result(1).announceDate='��׼��';
result(1).announceDatePrice='��׼�ռ۸�';
result(1).furtherDate='������';
result(1).actualPrice='�����ռ۸�';
result(1).furtherPrice='�����۸�';
result(1).dateInterval = '���ʱ��';
result(1).announceDateAccount='��׼���ۿ���';
result(1).actualToAnnouce='��Ʊ�����ʣ�δ�����г����أ�';
result(1).adjustedActualToAnnounceDate='��Ʊ������(����ڴ���)';

for i = 2 : size(data, 1)
    secu = data(i, :);
    result(i).code = secu{1};	% ��Ʊ��˾����
    result(i).name = secu{2};	% ��Ʊ��˾����

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
    actualPrice = securityPrice(secu{1}, furtherDate);
    if numel(actualPrice) ~= 1
        result(i).actualPrice = actualPrice{1, 8};    % ��������ʵ�ʼ۸�
    end

    result(i).announceDateAccount = secu{21};        % ��׼���ۿ���
    result(i).actualToAnnouce = result(i).actualPrice / ...
        result(i).announceDatePrice * 100;   % �ɼ�������

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
end

warning off MATLAB:xlswrite:AddSheet
xlswrite(filename, squeeze(struct2cell(result))', 'history');    % �����д��xls�ļ�