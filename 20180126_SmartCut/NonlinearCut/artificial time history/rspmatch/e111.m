close all

tetha = 0.1;
Sds=0.6;
T0D=1.6;
Sd1=Sds*T0D; 
Tall=[linspace(0.001,0.2,10) ,  linspace(0.2001,1.5,90)  ,  linspace(1.5001,5,50)   ];
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
T01 = T0D*Bs/B1;

% To=0.2*(0.78/0.6);                % 0.2Ts 較短週期 上升區      Ts=SD1/SDS
% Ts=0.78/0.6;                      % Ts  短周期 平台區
for ii=1:length(Tall)
    if(Tall(ii)<=0.2*T01)                                 % SDS(0.4+0.6*(T/0.2TS)) 
        targetall(ii)=Sds*( 0.4+(1/Bs-0.4)*Tall(ii)/(0.2*T01) );
    elseif(Tall(ii)<=T01 && Tall(ii)>0.2*T01)
        targetall(ii)=Sds/Bs;                          % SDS
    elseif (Tall(ii)<=2.5*T01 && Tall(ii)>T01)
        targetall(ii)=T0D*Sds/Tall(ii)/B1;              %下降區 SD1/T
    else
        targetall(ii)=0.4*Sds/Bs;              
    end
end

targetall=targetall';
plot(Tall,targetall,'b-')


tetha = 0.05;
Sds=0.6;
T0D=1.6;
Sd1=Sds*T0D; 
Tall=[linspace(0.001,0.2,10) ,  linspace(0.2001,1.5,90)  ,  linspace(1.5001,5,50)   ];
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

targetall=targetall';
hold on;
plot(Tall,targetall,'r-','linewidth',2)
hold on;
plot([T01 T01],[0 0.8],'b--')
hold on;
plot([2.5*T01 2.5*T01],[0 0.8],'b--')
hold on;
plot([T0 T0],[0 0.8],'m--')
hold on;
plot([2.5*T0 2.5*T0],[0 0.8],'m--')
ylim([0,0.8])
xlabel('Period(sec)','fontsize',16)
ylabel('S_a(g)','fontsize',16)
title('台北一區反應譜','fontsize',16)
legend('\xi = 20%','\xi = 5%')