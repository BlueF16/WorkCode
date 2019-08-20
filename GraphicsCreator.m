
%% Clear Workspace 
clc
clear all
close all
%% Read Excel
x=[0,0]
[num1,txt1,raw1] = xlsread ("Graphics_July_2018.xlsx");
[num2,txt2,raw2] = xlsread ("July.xlsx","Monthly values");
%% Create Bar Graph
eqdaynew=num2(1:2:end,1)
eqnightnew=num2(2:2:end,1)
eqdayold=[61.6;73.6;72.1;55.8;63.9;62.6;70.4;66.1;68.1;59.3]
eqnightold=[60.1;72.3;73.3;51.5;60.1;62.4;68.9;64.7;67.4;58.9]
for i=1:10
    
    x(end+1,1)=eqdaynew(i)
    x(end,2)=eqdayold(i)
    x(end,3)=eqnightnew(i)
    x(end,4)=eqnightold(i)
    
end
x(1,:)=[];


figure
clr = [0 0.8 0;
   0.3 0.8 0.8;
   0 0 1];
colormap(clr);
bar(x(1:5,1:4));
barvalues;
xticklabels({'NMS1','NMS2','NMS3','NMS4','NMS5'})%,'NMS6','NMS7','NMS8','NMS9','NMS10'})
legend('Day July 2019','Day July 2018','Night July 2019','Night July 2018');
title('LAeq monthly values during July 2019 and July 2018 - Part 1');
ylabel('dB');




figure
clr = [0 0.8 0;
   0.3 0.8 0.8;
   0 0 1];
colormap(clr);
bar(x(6:10,1:4));
barvalues;
xticklabels({'NMS6','NMS7','NMS8','NMS9','NMS10'})
legend('Day July 2019','Day July 2018','Night July 2019','Night July 2018');
title('LAeq monthly values during July 2019 and July 2018 - Part 2');
ylabel('dB');

%%
[status,sheets] = xlsfinfo("Graphics_July_2018.xlsx") ;
data={'NMS1Data.xlsx','NMS2Data.xlsx','NMS3Data.xlsx','NMS4Data.xlsx','NMS5Data.xlsx','NMS6Data.xlsx','NMS7Data.xlsx','NMS8Data.xlsx','NMS9Data.xlsx','NMS0Data.xlsx'}



for i=1:10

[num1,txt1,raw1] = xlsread ("Graphics_July_2018.xlsx",(sheets{i})) ;

[num2,txt2,raw2] = xlsread ((data{i}),"Daily Day");

oldmonthday=num1(2:2:end,1);
oldmonthnight=num1(3:2:end,1);
newmonthday= num2(2:end,3);
[num2,txt2,raw2] = xlsread ("NMS1Data.xlsx","Daily Night");
newmonthnight=num2(2:end,3);

clearvars num1 num2 txt1 txt2 raw1 raw2
nms=i
switch nms
    case 1
        dlimit=65;
        nlimit=55;
    case 2
        dlimit=60;
        nlimit=50;
    case 3
        dlimit=50;
        nlimit=40;
    case 4
        dlimit=65;
        nlimit=55;
    case 5
        dlimit=60;
        nlimit=50;
    case 6
        dlimit=50;
        nlimit=40;
    case 7
        dlimit=60;
        nlimit=50;
    case 8
        dlimit=65;
        nlimit=55;
    case 9
        dlimit=55;
        nlimit=45;
    case 10
        dlimit=60;
        nlimit=50;
end

%%
z=dlimit*ones(31,1);
u=nlimit*ones(31,1);
h(i)=figure('Name',sprintf('NMS%.0f Daytime',i))
x=linspace(1,31,31);
plot(x,newmonthday,'r',x,oldmonthday,'b','LineWidth',2)
hold on
xlim([1 31])
ylim([46 78])
xticks(linspace(1,31,16))
yticks(linspace(46,78,9))
title(sprintf('NMS%.0f Daytime Comparison',i))
xlabel('Date in July')
ylabel('Equivalent Noise Level dB(A)')
str = '#B76E79';
color = sscanf(str(2:end),'%2x%2x%2x',[1 3])/255;
plot(x,z,'--','Color',color,'LineWidth',1.5)
legend('LAeq Daytime 2019','LAeq Daytime 2018','Daytime Noise Limit')
hold off


%%
y(i)=figure('Name',sprintf('NMS%.0f Night time',i))
str = '#EDB120';
color = sscanf(str(2:end),'%2x%2x%2x',[1 3])/255;
plot(x,newmonthnight,'Color',color,'LineWidth',2) %,x,oldmonthnight,'LineWidth',2)
hold on
str = '#7E2F8E';
color = sscanf(str(2:end),'%2x%2x%2x',[1 3])/255;
plot(x,oldmonthnight,'Color',color,'LineWidth',2)
xlim([1 31])
ylim([46 78])
xticks(linspace(1,31,16))
yticks(linspace(46,78,9))
title(sprintf('NMS%.0f Night time Comparison',i))
xlabel('Date in July')
ylabel('Equivalent Noise Level dB(A)')
str = '#4DBEEE';
color = sscanf(str(2:end),'%2x%2x%2x',[1 3])/255;
plot(x,u,'--','Color',color,'LineWidth',1.5)

if nlimit<46
    legend('LAeq Night Time 2019','LAeq Night Time 2018',sprintf('Night Time Noise Limit %.0f db(A)',nlimit))
else
    legend('LAeq Night Time 2019','LAeq Night Time 2018','Night Time Noise Limit')
end


hold off
end
