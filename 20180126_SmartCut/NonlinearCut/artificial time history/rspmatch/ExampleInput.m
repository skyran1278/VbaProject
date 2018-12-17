%acc=load('El Centro imperial valley.txt');
% [time,accz,accn,acce]=textread('331_TAP001.txt');
% acc=accn/1000;

% [acc]=textread('El Centro imperial valley.txt');
% acc=acc;

dt=textread('CHICHI_TAP001.txt','%f',1,'headerlines',12);
groundTUNE=textread('CHICHI_TAP001.txt','%f','headerlines',11);
noscaleN=zeros(length(groundTUNE)/4,1);
noscaleE=zeros(length(groundTUNE)/4,1);
for i=1:length(groundTUNE)/4
    noscaleN(i)=groundTUNE(4*(i-1)+3);
    noscaleE(i)=groundTUNE(4*(i-1)+4);
end
noscaleN=noscaleN.*0.001; %cm/s^2 to m/s^2  �{���g���n��1000?
noscaleE=noscaleE.*0.001; 
acc=-noscaleN; %%%%%%%%%%%%%%%%%%%�n�_�V
 

%Period Subset and Target
Tall=[linspace(0.001,0.2,10) ,  linspace(0.2001,1.5,90)  ,  linspace(1.5001,5,50)   ];        %Input the entire period to be macthed in ascending order
Sds=0.6;
T0D=1.3;
Sd1=Sds*T0D; 
% To=0.2*(0.78/0.6);                % 0.2Ts ���u�g�� �W�ɰ�      Ts=SD1/SDS
% Ts=0.78/0.6;                      % Ts  �u�P�� ���x��
for ii=1:length(Tall)
    if(Tall(ii)<=0.2*T0D)                                 % SDS(0.4+0.6*(T/0.2TS)) 
        targetall(ii)=Sds*(0.4+3*Tall(ii)/T0D);
    elseif(Tall(ii)<=T0D && Tall(ii)>=0.2*T0D)
        targetall(ii)=Sds;                          % SDS
    elseif (Tall(ii)<=2.5*T0D && Tall(ii)>=T0D)
        targetall(ii)=Sd1/Tall(ii);              %�U���� SD1/T
    else
        targetall(ii)=0.4*Sds;              
    end
end
targetall=targetall';              %Input the target accordingly to the period being matched
