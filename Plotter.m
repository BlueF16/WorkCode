clc
clear all
close all
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
figure('Name',sprintf('NMS%.0f Daytime',i))
x=linspace(1,31,31);
plot(x,newmonthday,'r',x,oldmonthday,'b','LineWidth',2)
legend('LAeq Daytime 2019','LAeq Daytime 2018')
xlim([1 31])
ylim([46 78])
xticks(linspace(1,31,16))
yticks(linspace(46,78,9))
title(sprintf('NMS%.0f Daytime Comparison',i))
xlabel('Date in July')
ylabel('Equivalent Noise Level dB(A)')
yline(dlimit,'--','Day Limit')
figure('Name',sprintf('NMS%.0f Night time',i))


plot(x,newmonthnight,'Color','#EDB120','LineWidth',2) %,x,oldmonthnight,'LineWidth',2)
hold on
plot(x,oldmonthnight,'Color','#7E2F8E','LineWidth',2)
xlim([1 31])
ylim([46 78])
xticks(linspace(1,31,16))
yticks(linspace(46,78,9))
title(sprintf('NMS%.0f Night time Comparison',i))
xlabel('Date in July')
ylabel('Equivalent Noise Level dB(A)')
yline(nlimit,'--','Night Limit')

if nlimit<46
    legend('LAeq Night Time 2019','LAeq Night Time 2018',sprintf('Night Time Noise Limit %.0f db(A)',nlimit))
else
    legend('LAeq Night Time 2019','LAeq Night Time 2018')
end


hold off
end
