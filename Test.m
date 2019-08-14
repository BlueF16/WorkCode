rootfolder = 'C:\Users\mina_\Desktop\Work\August 4, 2019';   %the folder containing all the directories you want to process. (Assuming they're all in the same folder)
datafolder = {'AutoDownloadedFiles_SV_200_39721'};  %the folders to process within rootfolder
exportfolder='C:\Users\mina_\Desktop\Work\Exported';
for folderidx = 1:numel(datafolder)
    
    
    %% Clear Workspace and CMD
clc

clearvars -except rootfolder datafolder folderidx exportfolder
    
    currentfolder = fullfile(rootfolder, datafolder{folderidx});  %full path to folder
%% Intializing Data
data=zeros(1,15);
daydata=zeros(1,15);
nightdata=zeros(1,15);
timedata=cell(0,0);
td=cell(0,0);
tn=cell(0,0);
RemovedDay=NaT(1,1);


%% Building File Directory 

filelisting=dir(fullfile(currentfolder, '*.csv'));
number=length(filelisting);
%% Inputs
fprintf('Folder Path is %s \n',currentfolder);
nms=datafolder{folderidx};
nms=str2double(nms(end));

if nms==0
    nms=10;
end
fprintf('This Data belongs to NMS %.0f \n',nms);
formatSpec = 'The progam will now process %.0f files from \n %s - %s \n to \n %s - %s \nThis may take a long time(To end the program use Ctrl+C)\n';
fprintf(formatSpec,number,filelisting(1).name,filelisting(1).date,filelisting(number).name,filelisting(number).date);

%bol = input ('\n Is this one month of data? (Y(1)/N(2)): ');
bol=1;
t1=datetime(filelisting(1).date);
t2=datetime(filelisting(number).date);
t2=datevec(t2);
t2=[t2(1),t2(2)+1,1,0,0,0];
t2=datetime(t2);
comparedt=t1:minutes(1):t2;
comparedt=comparedt';
%% Limits
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
fprintf('NMS%.0f has been selected. My database shows that NMS%.0f limtis are : \n Day Limit: %.0f \n Night Limit: %.0f \n',nms,nms,dlimit,nlimit);


for i=1:number
%%Log Beginning of each Loop 
filename=filelisting(i).name;
fullfilename= fullfile(filelisting(i).folder, filelisting(i).name);
fprintf('\nLoop %.0f of %.0f, processing file %s',i,number,filename);

    
%%Import Data from CSV File / Find and Report Error if file not found

try
[num,txt,raw] = xlsread (fullfilename);
catch
   fprintf('\nError with this file %s, Moving to Next File',filename);
continue
end
%% Idenfity Type of Data either Night or Day


t=datetime(strcat(txt(22:6:end,2),txt(22:6:end,3)),'InputFormat','dd/MM/uuuu HH:mm:ss' );

h=hour(t(1));
m=month(t(1));


%% Take P1 Data
num(22:6:end, 15)=m;
data = vertcat(data, num(22:6:end, 1:15));
timedata=vertcat(timedata,t);
tempdata=num(22:6:end, 1:15);
x=size(tempdata,1);
for i=1:x

    
h=hour(t(i));
    
if (6<h) && (h<20)
    type='d';
else
    type='n';
end    
    
if type == 'd'
daydata = vertcat(daydata, tempdata(i,:)) ;
td=vertcat(td,t(i));
else
nightdata = vertcat(nightdata, tempdata(i,:));
tn=vertcat(tn,t(i));

end


end

    
end
    
%% Arrange Data, Find If Files have multiple months

%Remove Zeros
daydata=daydata(2:end,:);
nightdata=nightdata(2:end,:);
data=data(2:end,:);
% Find Missing Minutes
DateVector = datevec(timedata);
DateVector(:,6)=0;
DateVector=datetime(DateVector);
missing=setdiff(comparedt,DateVector);
% Remove Unrealistic Data
maxi=110;
mini=30;
indexdmax=find(daydata(:,2) > maxi);
while (size(indexdmax,1))>0

    
    daydata( indexdmax(1) , :) = [];
    RemovedDay(end+1)=td( indexdmax(1));
    td(indexdmax(1))=[];
    indexdmax=find(daydata(:,2) > maxi);
end


indexdmin=find(daydata(:,3)< mini);
while(size(indexdmin,1))>0

    
    daydata( indexdmin(1) , :) = [];
    RemovedDay(end+1)=td( indexdmin(1));
    td(indexdmin(1))=[];
    indexdmin=find(daydata(:,3)< mini);

end




indexnmax=find(nightdata(:,2) > maxi);

while (size(indexnmax,1))>0
 
    
    nightdata( indexnmax(1) , :) = [];
    RemovedDay(end+1)=tn( indexnmax(1));
    tn(indexnmax(1))=[];
    indexnmax=find(nightdata(:,2) > maxi);

end
indexnmin=find(nightdata(:,3) < mini);

while (size(indexnmin,1))>0

    
    nightdata( indexnmin(1) , :) = [];
    RemovedDay(end+1)=tn( indexnmin(1));
    tn(indexnmin(1))=[];
    indexnmin=find(nightdata(:,3) < mini);

end
RemovedDay(1,:)=[];
%Find Where the Change in Month Happens in the Data
monthchangeday=logical(diff(daydata(:,15)));
monthchangenight=logical(diff(nightdata(:,15)));
indexnight=find(monthchangenight);
indexday=find(monthchangeday);


in=sum(monthchangenight(:) == 1);
id=sum(monthchangeday(:) == 1);
if id==0
    [m,n]=size(daydata);
    rangehoursd(1,1)=1;
    rangehoursd(1,2)=m;
end
if in==0
    [m,n]=size(nightdata);
    rangenight(1,1)=1;
    rangenight(1,2)=m;
end

for i=1:id
[m,n]=size(daydata); 
rangehoursd(1,1)=1;

rangehoursd(i,2)=indexday(i);
rangehoursd(i+1,1)=indexday(i)+1;

rangehoursd(id+1,2)=m;
end


for i=1:in
[m,n]=size(nightdata); 
rangenight(1,1)=1;

rangenight(i,2)=indexnight(i);
rangenight(i+1,1)=indexnight(i)+1;

rangenight(in+1,2)=m;
end


in=in+1;
id=id+1;



for bol=1
id=1;
in=1;
m=size(daydata,1);
n=size(nightdata,1);
rangehoursd(1,1)=1;
rangehoursd(1,2)=m;
rangenight(1,1)=1;
rangenight(1,2)=n;


end
%% Calculating NPI, BGN , EVT
[m,n]=size(data);
for i=1:m

    data(i,16)= 0.2*(data(i,14) - 30); %BGN
    data(i,17)= 0.25 * ( data(i,4)- data(i,14) ); %EVT
    data(i,18)= data(i,16) + data (i,17); %NPI
end

[m,n]=size(daydata);
for i=1:m

    daydata(i,16)= 0.2*(daydata(i,14) - 30); %BGN
    daydata(i,17)= 0.25 * ( daydata(i,4)- daydata(i,14) ); %EVT
    daydata(i,18)= daydata(i,16) + daydata (i,17); %NPI
end

[m,n]=size(nightdata);
for i=1:m

    nightdata(i,16)= 0.2*(nightdata(i,14) - 30); %BGN
    nightdata(i,17)= 0.25 * ( nightdata(i,4)- nightdata(i,14) ); %EVT
    nightdata(i,18)= nightdata(i,16) + nightdata (i,17); %NPI
end




%% Calculating Monthly Data
for i=1:id
    a=rangehoursd(i,1);
    b=rangehoursd(i,2);
    m=b-a;
monthdaydata(i).Max= max(daydata(a:b,2));                           %Max
monthdaydata(i).Min= min(daydata(a:b,3));                           %Min
monthdaydata(i).LEQ=(log10((sum(10.^(daydata(a:b,4)./10)))/m))*10;  %LEQ
monthdaydata(i).LNN=(log10((sum(10.^(daydata(a:b,5)./10)))/m))*10;  %LNN
monthdaydata(i).L5=(log10((sum(10.^(daydata(a:b,7)./10)))/m))*10;  %L5
monthdaydata(i).L10=(log10((sum(10.^(daydata(a:b,8)./10)))/m))*10;  %L10
monthdaydata(i).L50=(log10((sum(10.^(daydata(a:b,11)./10)))/m))*10; %L50
monthdaydata(i).L90=(log10((sum(10.^(daydata(a:b,13)./10)))/m))*10; %L90
monthdaydata(i).L95=(log10((sum(10.^(daydata(a:b,14)./10)))/m))*10; %L95
monthdaydata(i).BGN=(log10((sum(10.^(daydata(a:b,16)./10)))/m))*10; %BGN
monthdaydata(i).EVT=(log10((sum(10.^(daydata(a:b,17)./10)))/m))*10; %EVT
monthdaydata(i).NPI=(log10((sum(10.^(daydata(a:b,18)./10)))/m))*10; %NPI
monthdaydata(i).From=td(a);%Dates
monthdaydata(i).To=td(b);%Dates
end


for i=1:in
    a=rangenight(i,1);
    b=rangenight(i,2);
    m=b-a;
monthnightdata(i).Max= max(nightdata(a:b,2));                           %Max
monthnightdata(i).Min= min(nightdata(a:b,3));                           %Min
monthnightdata(i).LEQ=(log10((sum(10.^(nightdata(a:b,4)./10)))/m))*10;  %LEQ
monthnightdata(i).LNN=(log10((sum(10.^(nightdata(a:b,5)./10)))/m))*10;  %LNN
monthnightdata(i).L5=(log10((sum(10.^(nightdata(a:b,7)./10)))/m))*10;   %L5
monthnightdata(i).L10=(log10((sum(10.^(nightdata(a:b,8)./10)))/m))*10;  %L10
monthnightdata(i).L50=(log10((sum(10.^(nightdata(a:b,11)./10)))/m))*10; %L50
monthnightdata(i).L90=(log10((sum(10.^(nightdata(a:b,13)./10)))/m))*10; %L90
monthnightdata(i).L95=(log10((sum(10.^(nightdata(a:b,14)./10)))/m))*10; %L95         
monthnightdata(i).BGN=(log10((sum(10.^(nightdata(a:b,16)./10)))/m))*10; %BGN
monthnightdata(i).EVT=(log10((sum(10.^(nightdata(a:b,17)./10)))/m))*10; %EVT
monthnightdata(i).NPI=(log10((sum(10.^(nightdata(a:b,18)./10)))/m))*10; %NPI
monthnightdata(i).From=tn(a);%Dates
monthnightdata(i).To=tn(b);%Dates
end

vars={'a','b','filename','formatSpec','h','i','id','in','m','monthchangeday','monthchangenight','n','num','indexday','indexnight','raw','type','txt'} ;
clear (vars{:}); 

%%  Hourly Data
tnhours=hour(td);
hourschange=logical(diff(tnhours));
indexhours=find(hourschange);

 ih=size(indexhours);
 for i=1:ih(1)
 [m,n]=size(tnhours); 
 rangehoursd(1,1)=1;
 
 rangehoursd(i,2)=indexhours(i);
 rangehoursd(i+1,1)=indexhours(i)+1;
 
 rangehoursd(ih(1)+1,2)=m;
 end

 tnhours=hour(tn);
hourschange=logical(diff(tnhours));
indexhours=find(hourschange);

 ih=size(indexhours);
 for i=1:ih(1)
 [m,n]=size(tnhours); 
 rangehoursn(1,1)=1;
 
 rangehoursn(i,2)=indexhours(i);
 rangehoursn(i+1,1)=indexhours(i)+1;
 
 rangehoursn(ih(1)+1,2)=m;
 end
 
 
 
 
 
 hours=hour(timedata);
hourschange=logical(diff(hours));
indexhours=find(hourschange);

 ih=size(indexhours);
 for i=1:ih(1)
 [m,n]=size(hours); 
 rangehours(1,1)=1;
 
 rangehours(i,2)=indexhours(i);
 rangehours(i+1,1)=indexhours(i)+1;
 
 rangehours(ih(1)+1,2)=m;
 end
 
  %% Ensures enough data in each range
  ih=size(rangehoursd,1);
  if rangehoursd(ih,1)==rangehoursd(ih,2)
     
     rangehoursd(ih-1,2)=rangehoursd(ih,2)
     rangehoursd(ih,:)=[];
 end
 
 
 
 ih=size(rangehoursn,1); 
 
 if rangehoursn(ih,1)==rangehoursn(ih,2)
     
     rangehoursn(ih-1,2)=rangehoursn(ih,2);
     rangehoursn(ih,:)=[];
 end
 
 
 
 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
ih=size(rangehoursd,1);
 for i=1:ih
    a=rangehoursd(i,1);
    b=rangehoursd(i,2);
    m=b-a;
hourlydaydata(i).Max= max(daydata(a:b,2));                           %Max
hourlydaydata(i).Min= min(daydata(a:b,3));                           %Min
hourlydaydata(i).LEQ=(log10((sum(10.^(daydata(a:b,4)./10)))/m))*10;  %LEQ
hourlydaydata(i).LNN=(log10((sum(10.^(daydata(a:b,5)./10)))/m))*10;  %LNN
hourlydaydata(i).L5=(log10((sum(10.^(daydata(a:b,7)./10)))/m))*10;   %L5
hourlydaydata(i).L10=(log10((sum(10.^(daydata(a:b,8)./10)))/m))*10;  %L10
hourlydaydata(i).L50=(log10((sum(10.^(daydata(a:b,11)./10)))/m))*10; %L50
hourlydaydata(i).L90=(log10((sum(10.^(daydata(a:b,13)./10)))/m))*10; %L90
hourlydaydata(i).L95=(log10((sum(10.^(daydata(a:b,14)./10)))/m))*10; %L95
hourlydaydata(i).BGN=(log10((sum(10.^(daydata(a:b,16)./10)))/m))*10; %BGN
hourlydaydata(i).EVT=(log10((sum(10.^(daydata(a:b,17)./10)))/m))*10; %EVT
hourlydaydata(i).NPI=(log10((sum(10.^(daydata(a:b,18)./10)))/m))*10; %NPI
hourlydaydata(i).From=td(a);%Dates
hourlydaydata(i).To=td(b);%Dates
end

ih=size(rangehoursn,1);
for i=1:ih
    a=rangehoursn(i,1);
    b=rangehoursn(i,2);
    m=b-a;
hourlynightdata(i).Max= max(nightdata(a:b,2));                           %Max
hourlynightdata(i).Min= min(nightdata(a:b,3));                           %Min
hourlynightdata(i).LEQ=(log10((sum(10.^(nightdata(a:b,4)./10)))/m))*10;  %LEQ
hourlynightdata(i).LNN=(log10((sum(10.^(nightdata(a:b,5)./10)))/m))*10;  %LNN
hourlynightdata(i).L5=(log10((sum(10.^(nightdata(a:b,7)./10)))/m))*10;   %L5
hourlynightdata(i).L10=(log10((sum(10.^(nightdata(a:b,8)./10)))/m))*10;  %L10
hourlynightdata(i).L50=(log10((sum(10.^(nightdata(a:b,11)./10)))/m))*10; %L50
hourlynightdata(i).L90=(log10((sum(10.^(nightdata(a:b,13)./10)))/m))*10; %L90
hourlynightdata(i).L95=(log10((sum(10.^(nightdata(a:b,14)./10)))/m))*10; %L95         
hourlynightdata(i).BGN=(log10((sum(10.^(nightdata(a:b,16)./10)))/m))*10; %BGN
hourlynightdata(i).EVT=(log10((sum(10.^(nightdata(a:b,17)./10)))/m))*10; %EVT
hourlynightdata(i).NPI=(log10((sum(10.^(nightdata(a:b,18)./10)))/m))*10; %NPI
hourlynightdata(i).From=tn(a);                                           %Dates
hourlynightdata(i).To=tn(b);                                             %Dates
end

%% Daily

tnday=day(td);
dayschange=logical(diff(tnday));
indexday=find(dayschange);

 ih=size(indexday,1);
 [m,n]=size(tnday); 
 for i=1:ih
 
 rangedaysd(1,1)=1;
 
 rangedaysd(i,2)=indexday(i);
 rangedaysd(i+1,1)=indexday(i)+1;
 
 rangedaysd(ih(1)+1,2)=m;
 
 
 end
 


 
 
if ih==0
    rangedaysd(1,1)=1
    rangedaysd(1,2)=m
end

 tnday=day(tn);
dayschange=logical(diff(tnday));
indexday=find(dayschange);

 ih=size(indexday,1);
 for i=1:ih
 [m,n]=size(tnday); 
 rangedayn(1,1)=1;
 
 rangedayn(i,2)=indexday(i);
 rangedayn(i+1,1)=indexday(i)+1;
 
 rangedayn(ih(1)+1,2)=m;
 end
 
 
 
 
 
 %% Ensures enough data in each range
 ih=size(rangedaysd,1);
  if rangedaysd(ih,1)==rangedaysd(ih,2)
     
     rangedaysd(ih-1,2)=rangedaysd(ih,2)
     rangedaysd(ih,:)=[];
 end
 
 
 
 ih=size(rangedayn,1); 
 
 if rangedayn(ih,1)==rangedayn(ih,2)
     
     rangedayn(ih-1,2)=rangedayn(ih,2);
     rangedayn(ih,:)=[];
 end
 
 
  
 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 
ih=size(rangedaysd,1);
 for i=1:ih
    a=rangedaysd(i,1);
    b=rangedaysd(i,2);
    m=b-a;
dailydaydata(i).Max= max(daydata(a:b,2));                           %Max
dailydaydata(i).Min= min(daydata(a:b,3));                           %Min
dailydaydata(i).LEQ=(log10((sum(10.^(daydata(a:b,4)./10)))/m))*10;  %LEQ
dailydaydata(i).LNN=(log10((sum(10.^(daydata(a:b,5)./10)))/m))*10;  %LNN
dailydaydata(i).L5=(log10((sum(10.^(daydata(a:b,7)./10)))/m))*10;   %L5
dailydaydata(i).L10=(log10((sum(10.^(daydata(a:b,8)./10)))/m))*10;  %L10
dailydaydata(i).L50=(log10((sum(10.^(daydata(a:b,11)./10)))/m))*10; %L50
dailydaydata(i).L90=(log10((sum(10.^(daydata(a:b,13)./10)))/m))*10; %L90
dailydaydata(i).L95=(log10((sum(10.^(daydata(a:b,14)./10)))/m))*10; %L95
dailydaydata(i).BGN=(log10((sum(10.^(daydata(a:b,16)./10)))/m))*10; %BGN
dailydaydata(i).EVT=(log10((sum(10.^(daydata(a:b,17)./10)))/m))*10; %EVT
dailydaydata(i).NPI=(log10((sum(10.^(daydata(a:b,18)./10)))/m))*10; %NPI
dailydaydata(i).From=td(a);%Dates
dailydaydata(i).To=td(b);%Dates
end

ih=size(rangedayn,1); 
for i=1:ih
    a=rangedayn(i,1);
    b=rangedayn(i,2);
    m=b-a;
dailynightdata(i).Max= max(nightdata(a:b,2));                           %Max
dailynightdata(i).Min= min(nightdata(a:b,3));                           %Min
dailynightdata(i).LEQ=(log10((sum(10.^(nightdata(a:b,4)./10)))/m))*10;  %LEQ
dailynightdata(i).LNN=(log10((sum(10.^(nightdata(a:b,5)./10)))/m))*10;  %LNN
dailynightdata(i).L5=(log10((sum(10.^(nightdata(a:b,7)./10)))/m))*10;   %L5
dailynightdata(i).L10=(log10((sum(10.^(nightdata(a:b,8)./10)))/m))*10;  %L10
dailynightdata(i).L50=(log10((sum(10.^(nightdata(a:b,11)./10)))/m))*10; %L50
dailynightdata(i).L90=(log10((sum(10.^(nightdata(a:b,13)./10)))/m))*10; %L90
dailynightdata(i).L95=(log10((sum(10.^(nightdata(a:b,14)./10)))/m))*10; %L95         
dailynightdata(i).BGN=(log10((sum(10.^(nightdata(a:b,16)./10)))/m))*10; %BGN
dailynightdata(i).EVT=(log10((sum(10.^(nightdata(a:b,17)./10)))/m))*10; %EVT
dailynightdata(i).NPI=(log10((sum(10.^(nightdata(a:b,18)./10)))/m))*10; %NPI
dailynightdata(i).From=tn(a);                                           %Dates
dailynightdata(i).To=tn(b);                                             %Dates

end


%%
if bol==1
    
    monthnightdata(1).NightLimit=nlimit;
    monthdaydata(1).DayLimit=dlimit;
    leqhourday=vertcat(hourlydaydata(:).LEQ);
    leqhournight=vertcat(hourlynightdata(:).LEQ);
    indday = find(leqhourday > dlimit);
    indnight = find(leqhournight > nlimit);
    monthdaydata(1).Exceedance=size(indday,1);
    monthnightdata(1).Exceedance=size(indnight,1);
    monthdaydata(1).Observations=size(rangehoursd,1);
    monthnightdata(1).Observations=size(rangehoursn,1);
    monthdaydata(1).Percent= size(indday,1)/ size(rangehoursd,1)*100;
    monthnightdata(1).Percent= size(indnight,1)/ size(rangehoursn,1)*100;
end
    



%% Exporting Data into XLS format
ExportHourDay=struct2table(hourlydaydata);
ExportHourNight=struct2table(hourlynightdata);
ExportDailyDay=struct2table(dailydaydata);
ExportDailyNight=struct2table(dailynightdata);
ExportMonthDay=struct2table(monthdaydata);
ExportMonthNight=struct2table(monthnightdata);
ExportFullData=[array2table(data) table(timedata)];
ExportFullData(:,15)=[];
ExportFullData.Properties.VariableNames={'TIME' 'MAX' 'MIN' 'LEQ' 'LNN' 'L1' 'L5' 'L10' 'L30' 'L40' 'L50' 'L60' 'L90' 'L95' 'BGN' 'EVT' 'NPI' 'Time'};
ExportMissing=table(missing);
Exportfilelisting=struct2table(filelisting);
ExportRemoved= table(RemovedDay);


filename = sprintf('NMS%.0fData.xlsx',nms);

writetable(ExportHourDay,filename,'Sheet','Hourly Day');
writetable(ExportHourNight,filename,'Sheet','Hourly Night');
writetable(ExportDailyDay,filename,'Sheet','Daily Day');
writetable(ExportDailyNight,filename,'Sheet','Daily Night');
writetable(ExportMonthDay,filename,'Sheet','Monthly Day');
writetable(ExportMonthNight,filename,'Sheet','Monthly Night');
writetable(ExportFullData,filename,'Sheet','Full Data');
writetable(Exportfilelisting,filename,'Sheet','Files Log');
writetable(ExportMissing,filename,'Sheet','Missing Files');
writetable(ExportRemoved,filename,'Sheet','Removed Minutes');

%% Delete Sheets 1,2,3

excelFileName = filename;
excelFilePath = pwd; % Current working directory.
sheetName = 'Sheet'; % EN: Sheet, DE: Tabelle, etc. (Lang. dependent)
% Open Excel file.
objExcel = actxserver('Excel.Application');
objExcel.Workbooks.Open(fullfile(excelFilePath, excelFileName)); % Full path is necessary!
% Delete sheets.
try
      % Throws an error if the sheets do not exist.
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '1']).Delete;
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '2']).Delete;
      objExcel.ActiveWorkbook.Worksheets.Item([sheetName '3']).Delete;
catch
      ; % Do nothing.
end
% Save, close and clean up.
objExcel.ActiveWorkbook.Save;
objExcel.ActiveWorkbook.Close;
objExcel.Quit;
objExcel.delete;

end