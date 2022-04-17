# ProcessSensorData-GEOPiles
Summary Description: Cleaning and Processing messy/missing sensor sata of ground heat exchange system - EbyRush Project

**function [outputArgs] = PreProcess_ExcelFiles(InputArgs)**

Author: Sylvie Antoun   
Date: 2022/04/14  
Email: sylvieantoun@gmail.com

PreProcess_ExcelFiles function lets you interactively handle messy and missing data values such as NaN or <missing>. 
The function main tasks are:
- Retime and Synchronize Time Dependent Sensor Variables.
- Find, fill, or remove missing data. 
 
Make sure you have downloaded the raw excel files for all sensors below from EbyRush - Argentum interface. To download the data, please refer to the "How To Document" located in EbyRush-Project shared drive.The sensor FileNames list is: [T1, T2, T3, T4, T5, T6, T7, TP1, TP2, TP3, TP4, TP5, TP6, TP7, TP8, F1, F2, F3, F4, EER (cooling) or COP (heating), EWT, LWT, HEHR, P1, P2, TotalCapacity, TotalPower, ZoneTempStat]
Make sure you change your Folder to the directory where your .xls* files exit. Please refer to the "How To Document" for more guidance 


**function [outputArgs] = PostProcess_CleanData(inputArgs)**

Author: Sylvie Antoun   
Date: 2022/04/14  
Email: sylvieantoun@gmail.com

1. PostProcess_CleanData function reads all the sensors Data and process it for evaluating and plotting Heat Exchange, Capacity, Power consumption and COP as a function of time for each EbyRush-System component:  [Existing Conventional, Borehole Loop, Geo-Piles,Heat Pump]

2. PostProcess_CleanData function is set up to assess both the COP of the Heat Pump (HP) and a Theoretical GeoPiles COP calculated by assuming the Heat Exchanger is absent. Theoretical COP is evaluated based on the Heat Pump Versatec VS Model 072 - Full Load Manual with T6 (water temperature coming out of the GeoPiles) used as the heat pump fluid temperature (EWT new) going directly into the Heat Pump. The interpolated data for the COP were based on a Heat Pump Fluid Flow Rate of 18 gpm. 

3. PostProcess_CleanData function is also set up to calculate the mean value of outputs over the test time. 

Make sure you change your Folder to the directory where your 'Date_Test_CleanData.xlsx file is saved. This Code is well-commented to guide through the steps if you need to make any changes. Please refer to the "How To Document" and "Technical Report" for more guidance.


**Extra Notes of Unit Conversion factors**
1 W = 3.412142 BTU/h, 1W= 0.000284345 ton, 1 ton = 12,000 BTU/h, and 1 ton = 3.51685 kW
F to degC =(F-32)/1.8, 
GPM to lps for flowrate lps=gpm*3.7854/60   1 gpm = 6.309e-5 m3/s
F1 is the flow rate of glycol flowing into the heat pump (gal/min), the same value as the heat pump WaterFlow variable. 
F2 and F3  are the flow rates of water going flowing into the geo-pile group1 and group2 loops respectively (gpm)
Cpwg is the specific heat of water with 30 glycol (485 Btu/Ib.F =3915 J/kg-K), 
Cpw is the specific heat of water (500 Btu/Ib.F =4187.00 J/kg-K), 
LWT is the leaving water temperature of the heat pump (F) 
EWT is the entering water temperature of the heat pump (F)
T7 is the temperature of the water entering the geo-piles, 
T6 is the temperature of the water leaving the entire geo-pile system, 
TotalHPCapacity and HEHR are in Btu/hr and Power in Watts ->  1 W = 3.412142 BTU/h
TotPowerHP is the total HP power consumption reading during operation (W)
P1 is the power consumed by pump 1 in figure 7.1 (W). 
P2 is the power consumed by pump 2 in figure 7.1 (W).


