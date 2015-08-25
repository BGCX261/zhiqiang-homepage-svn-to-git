function path = BrownMotionPath(t, para, pathNumber)
% �����˶�·��ģ��;���β����˶�·��ģ��;���ڲ����˶����������ģ��

if nargin <= 0
    t = [1, 0.01];
end
if nargin <= 1
    para = [0, 1];
end
if nargin <= 2
    pathNumber = 1;
end

if length(para) == 1
    para = [0, para];
end
if length(t) == 1
    t = [t, 0.01];
end

pointNumber = floor(t(1) / t(2));
path = zeros(pathNumber, pointNumber + 1);
for i = 1 : pathNumber
    path(i, 1) = para(1);
    for j = 1 : pointNumber + 2
        path(i, j + 1) = path(i, j) + sqrt(t(2)) * normrnd(0, 1);
    end
end

% %���������˶�·��,ʱ������(0,100)
% %�������β����˶�·��d(log(s))=mu*s*d(t)+sigma*s*d(w),w��ʾ�����˶�
% % price_stock() :
% mu=0.05;
% sigma=1;
% dt=1/1000;
% n_path=10;
% n_points=1001;
% price_stock=zeros(n_points,n_path);
% for j=1:n_path
%     price_stock(1,j)=1; %Ŀǰ�Ĺ�Ʊ�۸�Ϊ1
% end
% for j=1:n_path
%     for i=1:n_points-1
%         price_stock(i+1,j)=price_stock(i,j)*exp(dt*(mu-0.5*sigma^2)+sigma*sqrt(dt)*normrnd(0,1,1));
%     end
% end
% x=(0:0.001:1)';
% subplot(2,1,1);
% plot(x,brown_motion_values);
% xlabel('t')
% ylabel('y_value')
% title('�����˶�·��')
% subplot(2,1,2);
% plot(x, price_stock);
% xlabel('t')
% ylabel('y_value')
% title('���β����˶�·��')