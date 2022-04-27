function [outputArgs] = PreProcess_ExcelFiles(INPUTARGS)

% Author: Sylvie Antoun
% Date: 2022/04/14
% Version: 2.0
% Title: GEOPiles Eby-Rush Source Code
% Email: sylvieantoun@gmail.com
% Availability: https://github.com/sylvieantoun/ProcessSensorData-GEOPiles


% PREPROCESS_EXCELFILES function lets you interactively handle messy and missing data values such as NaN or <missing>. 
% The function main objectives is to: 
% $ Retime and Synchronize Time Dependent Sensor Variables.
% $ Find, fill, or remove missing data. 
% 
% [OUTPUTARGS] = PREPROCESS_EXCELFILES(INPUTARGS) 
% Make sure you have dowloaded the raw excel files for all sensors below from EbyRush-Argentum interface. 
% To download the data, please refer to the "How To Document" in EbyRush-Project Shared Drive
% The sensor FileNames will consist of the below:
% T1, T2, T3, T4, T5, T6, T7, TP1, TP2, TP3, TP4, TP5, TP6, TP7, TP8, F1, F2, F3, F4...
% EER (cooling) or COP (heating), EWT, LWT, HEHR, P1, P2, TotalCapacity, TotalPower, ZoneTempStat
% Make sure you change your Folder to the directory where your .xls* files exit
% Please refer to the "How To Document" for more guidance 

tic
clc
clear all

formatIn = 'yyyy-mm-dd HH:MM';
prompt = input('Enter Test Start Date and Time (format:yyyy-mm-dd HH:MM):','s')
timestartcar=char(prompt);
tstart= datetime(datevec(timestartcar(:,1:16),formatIn));
prompt = input('Enter Test End Date and Time (format:yyyy-mm-dd HH:MM):','s')
timeendcar=char(prompt);
tend= datetime(datevec(timeendcar(:,1:16),formatIn));

%% Importing Sensor Data
prompt = input('Please select all the sensors files, make sure you are in the right directory and your files are in .xlsx or .xls format')
[files,path]=uigetfile('*.xls*','multiselect','on');
% Get a list of all files in the folder with the desired file name pattern.
filePattern = fullfile(path, '*.xls*'); % Change to whatever pattern you need.
theFiles = struct2table(dir(filePattern));
ExcelFileNames=theFiles.name;
[folder, baseFileName, extension] = fileparts(ExcelFileNames);

numbGS=length(ExcelFileNames); % Total number of GroundLoop Sensors

% only process the data from t_start to t_finish 
for k = 1:numbGS
    opts{k}=detectImportOptions(ExcelFileNames{k});
    Data{k}=readtable(ExcelFileNames{k});
    VarName{k}=opts{k}.VariableNames;%variables Name Character
    timechar=char(Data{k}.time);%transfer the timedata into character 
    time{k}=datetime(datevec(timechar(:,1:16), formatIn));% only read timechar to seconds 1:16 char, TableArray
    indx=find(isbetween(time{k},tstart,tend)==1); %collect only the data between test time start and time end
    ttest{k}=time{k}(indx);
    TT{k}=timetable(ttest{k},str2double(Data{k}{indx,2}),'VariableNames',VarName{k}(1,2));
 if k== 1
        joinedtimetable = TT{k};  %1st time, create output
    else  
        joinedtimetable = synchronize(joinedtimetable, TT{k}); %subsequent times, synchronize. Choose whichever method is prefered
    end
end
 TT_all_unclean= synchronize(TT{:});

%% Ground loop snsors (28 total)
    TT_all_sort_irreg=sortrows(TT_all_unclean);

    if isregular(TT_all_sort_irreg)==0
       uniqueRowsTT=unique(TT_all_sort_irreg);
       dupTimes = sort(uniqueRowsTT.Time);
       tf = (diff(dupTimes) == 0);
       dupTimes = dupTimes(tf);
       dupTimes = unique(dupTimes); 
       uniqueTimes = unique(uniqueRowsTT.Time);
       firstUniqueRowsTT = retime(uniqueRowsTT,uniqueTimes,'firstvalue');
       meanTT = retime(uniqueRowsTT,uniqueTimes,'mean');
       regularTT = retime(meanTT,'regular','linear','TimeStep',minutes(1));
       TTclean= regularTT;
     else 
          TTclean= TT_all_unclean;
    end
TTclean = fillmissing(TTclean,'linear');    
%change timetable name based on test
formatSpec = "%s_Test_CleanData.xlsx";
testdate=timechar(1,1:10);
str = sprintf(formatSpec,testdate)
writetimetable(TTclean,str);
toc

end
