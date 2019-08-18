
%% Clear Workspace 
clc
clear all
close all
%% Read Excel
x=[0,0]
[num1,txt1,raw1] = xlsread ("Graphics_July_2018.xlsx");
[num2,txt2,raw2] = xlsread ("July.xlsx","Monthly values");
%%
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
