function [outputArgs] = PostProcess_CleanData(inputArgs)

%% PostProcess_CleanData function Summary
% Author: Dr. Sylvie Antoun | Date: 2022/04/14 | Email: sylvieantoun@gmail.com |

% [OUTPUTARGS] = POSTPROCESS_CLEANDATA(INPUTARGS) 
% The function reads all the sensors Data and process it for
% evaluating and plotting Heat Exchange, Capacity, Power consumption
% and COP as a function of time for each EbyRush-System component: 
% [Existing Conventional, Borehole Loop, Geo-Piles,Heat Pump]

% This function is setup to assess both the COP of the Heat Pump (HP) and a Theoretical GeoPiles COP calculated by assuming the Heat Exchanger is absent. 
% Theoretical COP is evaluated based on the Heat Pump Versatec VS Model 072 - Full Load Manual with T6 (water temperature coming out of the GeoPiles) 
% used as the heat pump fluid temperature (EWT new) going directly into the Heat Pump. 
% Note that the interpolated data for the COP were based on a Heat Pump Fluid Flow Rate of 18 gpm 

% This function is also setup to calculate the mean value of outputs over the test time. 

% Make sure you change your Folder to the directory where your 'Date_Test_CleanData.xlsx file is saved
% 
% This Code is well-commented to guide through the steps if you need to make any changes
% Please refer to the "How To Document" and "Technical Report" for more guidance 


%% Extra Notes of Unit Conversion factors 
% * 1 W = 3.412142 BTU/h, 1W= 0.000284345 ton, 1 ton = 12,000 BTU/h, and 1 ton = 3.51685 kW
% * F to degC =(F-32)/1.8, 
% * GPM to lps for flowrate lps=gpm*3.7854/60   1 gpm = 6.309e-5 m3/s
% * F1 is the flow rate of glycol flowing into the heat pump (gal/min), the same value as the heat pump WaterFlow variable.
% * F2 and F3  are the flow rates of water going flowing into the geo-pile group1 and group2 loops respectively (gpm)
% * Cpwg is the specific heat of water with 30% glycol (485 Btu/Ib.F =3915 J/kg-K), 
% * Cpw is the specific heat of water (500 Btu/Ib.F =4187.00 J/kg-K), 
% * LWT is the leaving water temperature of the heat pump (F) 
% * EWT is the entering water temperature of the heat pump (F)
% * T7 is the temperature of the water entering the geo-piles, 
% * T6 is the temperature of the water leaving the entire geo-pile system, 
% * TotalHPcapacity and HEHR are in Btu/hr and Power in Watts ->  1 W = 3.412142 BTU/h
% * TotPowerHP is the total HP power consumption reading during operation (W)
% * P1 is the power consumed by pump 1 in figure 7.1 (W). 
% * P2 is the power consumed by pump 2 in figure 7.1 (W). 

%% Importing Clean Data
clear
clc
%Antifreeze Correction Factor from Heat Pump manual 
AntifreezeCorrcooling=0.95;  AntifreezeCorrHeating =0.854;
formatIn = 'yyyy-mm-dd HH:MM';% time vector format

prompt = input('Make sure your current Folder is where you saved the Test Clean Data File. If true, Hit "Enter" to select the Clean file in .xlsx format')
[files,path]=uigetfile('*.xls*');
filePattern = fullfile(path, files); % Change to whatever pattern you need.
theFiles = struct2table(dir(filePattern));
ExcelFileNames=theFiles.name;
opts=detectImportOptions(ExcelFileNames);
%Read the data and variable names from file and save them in a Matlab time table
T=readtimetable(ExcelFileNames,opts,'ReadVariableNames',true);
VarName=opts.VariableNames; %variables Name Character

%%interpolate or remove rows with missing NaN values
T = fillmissing(T,'linear');
T = rmmissing(T,'DataVariables',@isnumeric);

if T.EWT < T.LWT %Cooling Season
    AntifreezeCorr=AntifreezeCorrcooling;
else 
     AntifreezeCorr=AntifreezeCorrHeating;
end 

%% Curve Fitting COP Heat Pump Specs using Versatec-VS Model 072 Full Load at 18 GPM	
%Values correspond to Flow = 18gpm
CF=readtable('Heat Pump Curve Data.xlsx','Sheet','Model72 - Flow 18','VariableNamingRule','preserve'); %reading data from excel
%F to degC =(F-32)/1.8, 
for i=1:size(CF.EWT)
  EWT_HPManual(i)=(CF.EWT(i)-32)/1.8;%convert to degC
  COP_Cooling_HPManual(i)=CF.EER(i)/3.412;
  COP_Heating_HPManual(i)=CF.COP(i);
end

if T.T6(10)<T.T7(10)
    x= EWT_HPManual;
    y= COP_Cooling_HPManual; 
else 
    x= EWT_HPManual;
    y= COP_Heating_HPManual;
end 
[xData, yData] = prepareCurveData( x, y );
% Set up fittype and options.
ft = fittype( 'poly3' );
% Fit model to data.
[fitresult, gof] = fit( xData, yData, ft );
% Plot fit with data.
figure0=figure;
axes0 = axes('Parent',figure0);
axes0.FontSize=12;
hold(axes0,'on');
box(axes0,'on');
plot( xData, yData,'ob','MarkerFaceColor','b','MarkerSize',4);
hold on
plot(fitresult,'predobs')
xlim([-5 50])
legend( axes0, 'Heat Pump Manual', 'CurveFit of COP(EWT)', 'Prediction Bounds','Location', 'Best' );
set(axes0,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],...
    'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
% Label axes
xlabel( 'EWT (degC)','FontSize', 12 );
ylabel( 'COP_{Cooling}', 'FontSize', 12 );
title('Versatec VS Model 072 - Full Load')
subtitle('COP Cooling - EAT 70 F (High Speed 2000 CFM) ', 'FontSize', 12')


%%Fluid Properties as a function of Fluid Temperature
Cp_w=1000*(28.07-0.2817*(T.T6+T.T7)/2+1.25*10^(-3)*((T.T6+T.T7)/2).^2 ...
    -2.48*10^(-6)*((T.T6+T.T7)/2).^3+1.857*10^(-9)*((T.T6+T.T7)/2).^4);%(T in Kelvin, unit is J/kg.K)
Cp_w=4187;
Cp_wg=3915;
rho_w=1001.1-0.0867*(T.T6+T.T7)/2-0.0035*((T.T6+T.T7)/2).^2;

Twg=[-13	-7	-3	0	20	40	60	80	100];
rho_wg=[1035	1034	1032	1031	1022	1010	997	983	969];
[xData, yData] = prepareCurveData( Twg, rho_wg );
% Set up fittype and options.
ft1 = fittype( 'poly3' );
% Fit model to data.
[rho_wg, gof1] = fit( xData, yData, ft1 );

a=-7.75E-09;
b=1.03E-06;
c=-5.71E-05;
d=1.79E-03;

Do=0.15;
mu_w2 = a*((T.T6+T.T7)/2).^3 + b*((T.T6+T.T7)/2).^2 + c*((T.T6+T.T7)/2) + d;%[Pa s], [N s/m2]Do=0.1143;%m
Ac=(pi*Do^2)/4;%m2
vi= ((T.F2+T.F3)*6.309e-5/8)/Ac; %velocity per pile m/s
Re=rho_w.*Do.*vi./mu_w2;

%%Post-Calculations and Saving data in  'Out' structure 
Out.Time=T.Time;
Out.HEHR=T.HEHR/3.412142/1000;%Btu/hr to kW
Out.HE_HP=abs((T.T2-T.T1).*T.F1.*6.309e-5*Cp_wg.*rho_wg((T.T1+T.T2)/2))/1000;% deltaT[K]*flow(gpm)*6.309e-5[m3/s/gpm]*cp[J/kg-K]*rho[kg/m3]/1000 (kW)
%At Heat Exchanger Secondary Fluid in Pipe going into ground
Out.HE_geopile_sf=abs((T.T6-T.T7).*(T.F3+T.F2).*6.309e-5.*Cp_w.*rho_w/1000); % deltaT[K]*flow(gpm)*6.309e-5[m3/s/gpm]*cp[J/kg-K]*rho[kg/m3]/1000 (kW)
Out.HE_conv=abs((T.T4-T.T5).*T.F4*6.309e-5*Cp_wg.*rho_wg((T.T4+T.T5)/2))/1000;

% At Heat Exchanger Primary fluid in Pipe going to Heat Pump
%Out.HE_geopile_pf=Out.HE_HP-Out.HE_conv;
Out.deltaT_geopiles=abs(T.T6-T.T7);
%Out.deltaT_HeatExchanger=abs(Out.HE_geopile_sf*1000/(Cp_wg*rho_w.*(T.F1-T.F4)));

% Capacities in Tons
Out.TotalCapacity_HP=T.TotalCapacity/12000;%in tons  12000 Btu/hr=1 ton
Out.PartialCapacity_piles=Out.TotalCapacity_HP.*(T.F1-T.F4)./T.F1;%in tons
Out.PartialCapacity_Conv=Out.TotalCapacity_HP-Out.PartialCapacity_piles;%in tons

%Power in kW
if ismember('P1',T.Properties.VariableNames)==0
T.P1=T.CT1PowerConsumption;
else 
    T.P1=T.P1;
end

if ismember('CT2PowerConsumption',T.Properties.VariableNames)==0 && ismember('P2',T.Properties.VariableNames)==0
T.P2=ones(length(T.T6),1)*10;
end
if ismember('P2',T.Properties.VariableNames)==0 && ismember('CT2PowerConsumption',T.Properties.VariableNames)==1
   T.P2=T.CT2PowerConsumption;
end

Out.TotalPower_HP_includingP1P2=(T.TotalPower+T.P1+T.P2)/1000;%kW
Out.PartialPower_piles=(T.F1-T.F4)./T.F1.*Out.TotalPower_HP_includingP1P2;%kW
Out.PartialPower_Conv=Out.TotalPower_HP_includingP1P2-Out.PartialPower_piles;%kW

%% COP Calculation
Out.COP_HP_includingP1P2= (Out.TotalCapacity_HP*3.51685)./Out.TotalPower_HP_includingP1P2;%kW/kW convert capacity 1 ton = 3.51685 kW
% Out.COP_geopiles= (Out.PartialCapacity_piles*3.51685)./(Out.PartialPower_piles);%kW/kW
EER_HPreading=T.TotalCapacity./T.TotalPower;
Out.COP_HPreading=EER_HPreading/3.412;%or from HP EER readings COP = EER/3.412
% if T.T6 < T.T7 %Cooling 
Out.COP_TheoreticalGeoPiles= fitresult(T.T6);

%% Mean Values

Mean.HeatExchangeHP_kW=mean(Out.HE_HP);
Mean.HeatExchangeGeopiles_kW=mean(Out.HE_geopile_sf);
Mean.HeatExchangeConvLoop_kW=mean(Out.HE_conv);

Mean.TotalHPPowerConsump_kW=mean(Out.TotalPower_HP_includingP1P2);%in kW
Mean.GeoPilesPowerConsump_kW=mean(Out.PartialPower_piles);%in kW
Mean.ConvLoopPowerConsump_kW=mean(Out.PartialPower_Conv);%in kW

Mean.TotalCapacityHP_Tons=mean(Out.TotalCapacity_HP);%in tons
Mean.GeoPilesCapacity_Tons=mean(Out.PartialCapacity_piles);%in tons
Mean.ConvLoopCapacity_Tons=mean(Out.PartialCapacity_Conv);%in tons


Mean.HeatPumpFlowrate_GPM=mean(T.F1);%in GPM
Mean.GeopilesHXFlowrate_GPM=mean(T.F1)-mean(T.F4);
Mean.ConvLoopFlowrate_GPM=mean(T.F4);%in GPM

Mean.deltaTempPiles_degC= mean(Out.deltaT_geopiles);
Mean.deltaTempHP_degC= mean((abs(T.T1-T.T2)));
Mean.deltaTempConLoop_degC= mean((abs(T.T4-T.T5)));

Mean.COP_HP=mean(Out.COP_HP_includingP1P2);%
Mean.COP_Theoretical_Geopiles=mean(Out.COP_TheoreticalGeoPiles);

%% Plotting results 
% You can add x-axis time limits if you want to avoid transition time at start and end of the test
% example add a line "xlim([T.Time(1)+ minutes(10) T.Time(end)-minutes(5)]" to each figure if necessary; 

figure21=figure;
axes21 = axes('Parent',figure21);
axes21.FontSize=12;
hold(axes21,'on');
box(axes21,'on');
yyaxis left
plot(T.Time,Out.COP_TheoreticalGeoPiles,'LineWidth',1.5)
xlabel('datetime','FontSize', 12')
ylabel ('COP_{Theoretical-GeoPiles} ', 'FontSize',12)
yyaxis right
plot(T.Time,T.T6,'LineWidth',1.5)
ylabel ('Temp out of Geopiles (\circC) ', 'FontSize',12)
set(axes21,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],...
    'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('COP_{Theoretical GeoPiles}', 'Temp out of the piles','FontSize',10,'Location','Best')
title({'COP Theoretical of the GeoPiles Based on Heat Pump Manual' ,'Assuming T6 to be the EWT '}, 'FontSize', 12')

figure1=figure;
axes1 = axes('Parent',figure1);
hold(axes1,'on');
box(axes1,'on');
Markers = {'*','x','o','d','v','s','+'};
col={'k','r',[0.2 0.1 0.5],[0.1 0.5 0.8], [0.2 0.7 0.6],[ 0.8 0.7 0.3],[0.9 1 0]};
% plot(T.Time,(T.LWT-32)/1.8 -(T.EWT-32)/1.8 ,'LineWidth',1.5);
% hold on
plot(T.Time,T.T2-T.T1 ,'LineWidth',1.5);hold on ;
plot(T.Time,T.T4-T.T5 ,'LineWidth',1.5);hold on ;
plot(T.Time,T.T6-T.T7 ,'LineWidth',1.5);
% xlim([T.Time(1) T.Time(end-10)])
xlabel('datetime','FontSize',14)
ylabel ('Temperature (degC)', 'FontSize',14)
set(axes1,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],...
    'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('across HP','across Existing Conv Loop','across Geopiles','FontSize',14,'Location','Best')
title('Temp Diff across HP, Existing Loop and Geopiles', 'FontSize', 14')

figure1a=figure;
axes1a = axes('Parent',figure1a);
hold(axes1a,'on');
box(axes1a,'on');
plot(T.Time,T.T2 ,'LineWidth',1.5);
hold on;
plot(T.Time,T.T1,'LineWidth',1.5);
xlabel('datetime','FontSize',14)
ylabel ('Temperature (degC)', 'FontSize',14)
set(axes1a,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],...
    'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('Temp into the HP','Temp out of the HP','FontSize',14,'Location','Best')
title('T1 and T2 at HP', 'FontSize', 14')

figure2=figure;
axes2 = axes('Parent',figure2);
axes2.FontSize=14;
hold(axes2,'on');
box(axes2,'on');
plot(T.Time,T.T6,'LineWidth',1.5)
hold on;
plot(T.Time,T.T7,'LineWidth',1.5)
xlabel('datetime','FontSize',14)
ylabel ('Temperature (degC)', 'FontSize',14)
set(axes2,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],...
    'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('Temp out of the piles','Temp into the piles','FontSize',14,'Location','Best')
title('T7 and T6 Geopiles ', 'FontSize', 14')

figure2a=figure;
axes2a = axes('Parent',figure2a);
axes2a.FontSize=14;
hold(axes2a,'on');
box(axes2a,'on');
plot(T.Time,T.T4,'LineWidth',1.5)
hold on;
plot(T.Time,T.T5,'LineWidth',1.5)
xlabel('datetime','FontSize',14)
ylabel ('Temperature (degC)', 'FontSize',14)
set(axes2a,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],...
    'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('Temp out of the ConvLoop','Temp into the ConvLoop','FontSize',14,'Location','Best')
title('T4 and T5 Conventional Loop ', 'FontSize', 14)

figure3=figure;
axes3 = axes('Parent',figure3);
hold(axes3,'on');
box(axes3,'on');
plot(T.Time,Out.HE_HP,'LineWidth',1.5);hold on;
plot(T.Time,Out.HE_geopile_sf,'LineWidth',1.5);hold on;
plot(T.Time,Out.HE_conv,'LineWidth',1.5);hold on 
%xlim([T.Time(1) T.Time(end-2)])
xlabel('datetime','FontSize', 14')
ylabel('Heat Exchange(kW)', 'FontSize', 14');
set(axes3,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('HE_{HP}','HE_{Geopiles}','HE_{ConvLoop}', 'FontSize',14,'Location','Best');
title('Heat Exchange[kW]', 'FontSize', 14);
% 1 J/s= 3.412142 BTU/h

figure4=figure;
axes4 = axes('Parent',figure4);
hold(axes4,'on');
box(axes4,'on');
plot(T.Time,Out.TotalCapacity_HP,'LineWidth',1.5);hold on;
plot(T.Time,Out.PartialCapacity_piles,'LineWidth',1.5);hold on;
plot(T.Time,Out.PartialCapacity_Conv,'LineWidth',1.5);
xlabel('datetime','FontSize', 14')
ylabel('Capacity (Tons)', 'FontSize', 14);
set(axes4,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('TotalCapacity_{HP}','PartialCapacity_{piles}','PartialCapacity_{conv}', 'FontSize',14,'Location','Best');
title('Capacities [Tons]', 'FontSize', 14);
% 1 J/s= 3.412142 BTU/h

figure5=figure;
axes5 = axes('Parent',figure5);
hold(axes5,'on');
box(axes5,'on');
plot(T.Time,T.F1,'LineWidth',1.5);hold on;
plot(T.Time,T.F1-T.F4,'LineWidth',1.5);hold on;
plot(T.Time,T.F4,'LineWidth',1.5);
xlabel('datetime','FontSize', 14')
ylabel('flowrate (gpm)', 'FontSize', 14');
set(axes5,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('Total HP Flow','PartialFlow_{HX}','PartialFlow_{conv}', 'FontSize',14,'Location','Best');
title('Flowrates [gpm]', 'FontSize', 14');
% 1 J/s= 3.412142 BTU/h

figure6=figure;
axes6 = axes('Parent',figure6);
hold(axes6,'on');
box(axes6,'on');
plot(T.Time,T.TP1,'LineWidth',1.5);hold on;plot(T.Time,T.TP2,'LineWidth',1.5);hold on;...
plot(T.Time,T.TP3,'LineWidth',1.5);hold on;plot(T.Time,T.TP4,'LineWidth',1.5);hold on;...
plot(T.Time,T.TP5,'LineWidth',1.5);hold on; plot(T.Time,T.TP6,'LineWidth',1.5);hold on;
%plot(T.Time,T.TP7,'LineWidth',1.5); hold on;  %% if Sensor TP7 working uncomment plot
%plot(T.Time,T.TP8,'k','LineWidth',1.5); %% if Sensor TP8 working uncomment plot
xlabel('datetime','FontSize', 14)
ylabel('Pile Outlet Temperature (C)', 'FontSize', 14');
set(axes6,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('TP1', 'TP2', 'TP3', 'TP4', 'TP5', 'TP6', 'TP7','FontSize',14,'Location','Best')
title('Pile Outlet Temperature (C)', 'FontSize', 14);

figure7=figure;
axes7 = axes('Parent',figure7);
axes7.FontSize=14;
hold(axes7,'on');
box(axes7,'on');
plot(T.Time,Out.TotalPower_HP_includingP1P2,'LineWidth',1.5);hold on;
plot(T.Time,Out.PartialPower_piles,'LineWidth',1.5);hold on;
plot(T.Time,Out.PartialPower_Conv,'LineWidth',1.5);
xlabel('datetime','FontSize', 14')
ylabel('PowerConsumption (kW)', 'FontSize', 14');
set(axes7,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('TotalPower_{HP}','PartialPower_{piles}','PartialPower_{conv}', 'FontSize',14,'Location','Best');
title('Total Power Consumption including P1 and P2[kW]', 'FontSize', 14');
% 1 J/s= 3.412142 BTU/h

figure8a=figure;
axes8a = axes('Parent',figure8a);
axes8a.FontSize=14;
hold(axes8a,'on');
box(axes8a,'on');
% plot(T.Time,EER_manuf/3.412,'LineWidth',1.5);hold on;
plot(T.Time,Out.COP_HP_includingP1P2,'LineWidth',1.5);hold on;
plot(T.Time,Out.COP_HPreading,'LineWidth',1.5);hold on;
%  plot(T.Time(10:end),COP_TheoreticalGeoPiles(10:end),'LineWidth',1.5);
xlim([T.Time(1)+minutes(5) T.Time(end)])
xlabel('datetime','FontSize', 14')
ylabel('COP', 'FontSize', 14');
set(axes8a,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('COP_{HP including P1 and P2}', 'COP_{HP excluding P1 and P2}','FontSize',14,'Location','Best');
title('Heat Pump COP', 'FontSize', 14);

figure9=figure;
axes9 = axes('Parent',figure9);
hold(axes9,'on');
box(axes9,'on');
plot(T.Time,T.F2,'LineWidth',1.5);hold on;
plot(T.Time,T.F3,'LineWidth',1.5);hold on;
plot(T.Time,T.F2+T.F3,'LineWidth',1.5);
xlabel('datetime','FontSize', 14')
ylabel('flowrate (gpm)', 'FontSize', 14');
set(axes9,'GridColor',[0 0 0],'MinorGridColor',[0 0 0],'XColor',[0 0 0],'XMinorTick','on','YColor',[0 0 0],...
    'YMinorTick','on','ZColor',[0 0 0],'TickLength',[0.02 0.025]);
legend('Piles Group2 Flow F2','Piles Group1 Flow F3 ','Total Piles Flow (F2+F3)', 'FontSize',12,'Location','Best');
title('Flowrates [gpm]', 'FontSize', 14');

end
