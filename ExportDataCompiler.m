clc
clear all
%% Input
exportfolder='C:\Users\mina_\Desktop\Work\NMS Monthly\August\Exported'; %folder to export the collected data
filename= input('Export Data Filename(Append with .xlsx): ','s')
filelisting=dir (fullfile(exportfolder, '*NMS*.xlsx'));
fullfilename= fullfile(filelisting(2).folder, filelisting(2).name);
template=fullfile(exportfolder,'MonthlyTemplate.xlsx');
filename=fullfile(exportfolder, filename);

%% This program retrieves all exported excel sheets by CSVParser and compiles them into one excel sheet.
for i=1:length(filelisting)
    
    fullfilename= fullfile(filelisting(i).folder, filelisting(i).name);
    [num,txt,raw] = xlsread(fullfilename,'Hourly Data');
    
    

    data.DateTime = txt(2:end,13);
    data.Laeq=num(:,3);
    data.Lmax=num(:,1);
    data.Lmin=num(:,2);
    data.L5=num(:,5);
    data.L10=num(:,6);
    data.L50=num(:,7);
    data.L90=num(:,8);
    data.L95=num(:,9);
    data.NPI=num(:,12);
    data.BGN=num(:,10);
    data.EVT=num(:,11);
    T=struct2table(data);

    sheet=str2double (filelisting(2).name(4))
    sheetname= sprintf('NMS%.0f',sheet);
    titledate=strcat(data.DateTime(1),' to  ',data.DateTime(end));
    
    copyfile(template,filename);
    writetable(T,filename,'Sheet',sheetname,'Range','A10');
    xlrange='A8';
    xlswrite(filename,titledate,sheetname,xlrange);
    %% Monthly Tab
    [num,txt,raw] = xlsread(fullfilename,'Monthly Day');
    [num2,txt2,raw2] = xlsread(fullfilename,'Monthly Night');
    daymonthdata.Laeq=num(:,3);
    daymonthdata.Lmax=num(:,1);
    daymonthdata.Lmin=num(:,2);
    daymonthdata.L5=num(:,5);
    daymonthdata.L10=num(:,6);
    daymonthdata.L50=num(:,7);
    daymonthdata.L90=num(:,8);
    daymonthdata.L95=num(:,9);
    daymonthdata.NPI=num(:,12);
    daymonthdata.BGN=num(:,10);
    daymonthdata.EVT=num(:,11);
    
    
    
    nightmonthdata.Laeq=num2(:,3);
    nightmonthdata.Lmax=num2(:,1);
    nightmonthdata.Lmin=num2(:,2);
    nightmonthdata.L5=num2(:,5);
    nightmonthdata.L10=num2(:,6);
    nightmonthdata.L50=num2(:,7);
    nightmonthdata.L90=num2(:,8);
    nightmonthdata.L95=num2(:,9);
    nightmonthdata.NPI=num2(:,12);
    nightmonthdata.BGN=num2(:,10);
    nightmonthdata.EVT=num2(:,11);
    
    
    switch sheet
        
        case 0
            daystart='C28';
            nightstart='C29';
        case 1
            daystart='C10';
            nightstart='C11';
        case 2
            daystart='C12';
            nightstart='C13';
        case 3
            daystart='C14';
            nightstart='C15';      
         case 4
            daystart='C16';
            nightstart='C17';
         case 5
            daystart='C18';
            nightstart='C19';       
         case 6
            daystart='C20';
            nightstart='C21';
        case 7
            daystart='C22';
            nightstart='C23';
        case 8
            daystart='C24';
            nightstart='C25';
        case 9
            daystart='C26';
            nightstart='C27';
    end   
            
            
            
            
            
daym=struct2table(daymonthdata);
nightm=struct2table(nightmonthdata);
writetable(daym,filename,'Sheet','Monthly values','Range',daystart,'WriteVariableNames',0);
writetable(nightm,filename,'Sheet','Monthly values','Range',nightstart,'WriteVariableNames',0);
xlrange='A7';
xlswrite(filename,titledate,'Monthly values',xlrange);        
            
            
            
            
            
            
            
            
            
            
            
end