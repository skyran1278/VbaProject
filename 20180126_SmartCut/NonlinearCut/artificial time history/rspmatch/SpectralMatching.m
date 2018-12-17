close all;
clear all;
clc;
tic;
%--------------------------------------------------------------------------
%ver1.0 Last update on: 26/4/2013
%This routine performs spectral matching using the Corrected Tapered Cosine 
%Wavelet coupled with Broyden Updating. The spectra response is the
%pseudo-acceleration spectra. The elastic SDOF response is calculated using
%Newmark Method

%The steps for using this routine are as follow:
    %1. Input the required information on the General Input, including:
    %   - the vector 'acc' which contains the acceleration time series (g)
    %   - the time increment of acceleration time series, named as 'dt'
    %   - the damping level of the SDOF oscillator, named as 'tetha'
    %   - the tolerance limit of average misfit (g) and maximum error, named 
    %     as 'avgtol'and 'errortol' respectively
    %   - the zero pad duration (sec) at ends of the acceleration time series,
    %     named as 'zeropad' 
    
    %2. Input the required information on the Period Subset(s) and Target,
    %   including:
    %  - the period being matched (sec) in ascending order, named as 'Tall'  
    %  - the target response spectra (g) of the corresponding periods,named 
    %    as 'targetall'
    %  - the period ranges of each period subset, named as T1range, T2range
    %    and T3range
    
    %3. Input the required information on the gain factors
    
    %4. Run the Analysis (F5)
%--------------------------------------------------------------------------

% Load Example Input
% ExampleInput; %台北二區!!!!!!!!!!!!!!
% dat = importdata('El Centro.txt');
data1 = textread('chi chi_TAP024.txt','','headerlines',11);
data2 = textread('chi chi_TAP089.txt','','headerlines',11);
data3 = textread('chi chi_TAP067.txt','','headerlines',11);
data4 = textread('chi chi_TCU050.txt','','headerlines',11);

% Input the acceleration time series in g
% General Input
acc = data4(:,4)/981; 
t = data4(:,1);
dt = t(2) - t(1);   %dt=0.02;         %Ground motion time step (sec)
tetha=0.05;             %Damping level
avgtol=0.002;           %Tolerance on average misfit
errortol=0.0015;        %Tolerance on maximum error, 0.1 equals to 10%
zeropad= 1;              %Length of zero pad at the beginningg and end of the ground motion (sec)

%Period Subset and Target
Tall=[linspace(0.001,0.2,10) ,  linspace(0.2001,1.5,90)  ,  linspace(1.5001,5,50)   ];  %Input the entire period to be macthed in ascending order
T1range=[0.07 1.0];     %Input the minimum and maximum period in Period subset 1. Example of valid use=[0.05 1.5]
T2range=[0.07 3.0];     %Input the minimum and maximum period in Period subset 2  Example of valid use=[0.05 3.0], use 0 if no period subset 2
T3range=[0.05 5];            %Input the minimum and maximum period in Period subset 3, Example of valid use=[0.05 6.0], use 0 if no period subset 3
T4range=[0.05 5]; 
%Period Subset and Target
Sds=0.7;
T0D=4/7;
Sd1=Sds*T0D; 
switch tetha
    case 0.05
        Bs= 1;
        B1= 1;
    case 0.15
        Bs= 1.465;
        B1= 1.375;
    case 0.1
        Bs= 1.33;
        B1= 1.25;
    case 0.2
        Bs= 1.6;
        B1= 1.5;
end
T0 = T0D*Bs/B1;


% To=0.2*(0.78/0.6);                % 0.2Ts 較短週期 上升區      Ts=SD1/SDS
% Ts=0.78/0.6;                      % Ts  短周期 平台區
for ii=1:length(Tall)
    if(Tall(ii)<=0.2*T0)                                 % SDS(0.4+0.6*(T/0.2TS))
        targetall(ii)=Sds*( 0.4+(1/Bs-0.4)*Tall(ii)/(0.2*T0) );
    elseif(Tall(ii)<=T0 && Tall(ii)>0.2*T0)
        targetall(ii)=Sds/Bs;                          % SDS
    elseif (Tall(ii)<=2.5*T0 && Tall(ii)>T0)
        targetall(ii)=T0D*Sds/Tall(ii)/B1;              %下降區 SD1/T
    else
        targetall(ii)=0.4*Sds/Bs;
    end
end
targetall=targetall';              %Input the target accordingly to the period being matched
%Gain Factors
Cgain=1e2;              %Gain factor on initial C
wavmag=1e-7;            %Wavelet magnitude, g  
CoffDiag=0.7;           %Coefficient of the off-diagonal terms  
dbgain=1;               %Gain factor on vector b updating
dCgain=1;               %Gain factor on C updating
Broydeniter=10;         %Maximum number of Broyden iteration
Outeriter=6;            %Maximum number of Outer-loop iteration


%Execution
[amod]=SpectralMatchingFun(acc,dt,tetha,avgtol,errortol,zeropad,Cgain,wavmag,CoffDiag,dbgain,dCgain,...
    Broydeniter,Outeriter,Tall,T1range,T2range,T3range,T4range,targetall);


toc;
