function [abest]=SpectralMatchingFun(arec,dt,tetha,avgtol,errortol,zeropad,Cgain,wavmag,CoffDiag,dbgain,dCgain,...
    Broydeniter,Outeriter,Tall,T1range,T2range,T3range,T4range,targetall)

%This routine perfoms the spectral matching using wavelets and Broyden
%Updating.
%Available results are:
 %  - Spectrum compatible acceleration time series (abest)
 %  - Modified Response Spectrum (R)
 %  - Original and Spectrum compatible velocity time series (voriginal and vmod) 
 %  - Original Spectrum compatible displacement time series (doriginal and dmod)
 %  - Original and Modified Arias Intensity (IAori and IAproposed)
  
 
%Properties
m=1;                        %mass
iter=1;                     %starting iteration count
wall=2*pi./Tall;

%Zero padding at start and end of acceleration
[arec,t]=Zeropad(zeropad,arec,dt);
abest=arec; amod=arec;

%Form the period subset
[T1,T2,T3,T4,target1,target2,target3,target4]=PeriodSubset(Tall,T1range,T2range,T3range,T4range,targetall);

%initial value of best misfit for all targets
bestmisfitall=1000;
for OuterLoop=1:Outeriter
    
%PERIOD SUBSET 1
w=2*pi./T1; %angular frequency

%Step 1, calculate initial response
[u,t_peak,t_index,apeak,misfit,target]=IniResponse(amod,t,m,tetha,dt,T1,w,target1);

%Step 2, Least square scalling
[amod]=AccScale(apeak,target,amod);

%Step 3: Scale the acceleration time series based on the least square fit
[u,t_peak,t_index,apeak,misfit]=IniResponse(amod,t,m,tetha,dt,T1,w,target1);

%Step 4, calculate initial b, initial C and initial misfit
[b,C,misfit,wavtot]=InitialC2(T1,w,m,tetha,dt,t_peak,t_index,t,Cgain,target1,misfit,amod,CoffDiag,wavmag);

%Step 5, Broyden Loop
bestmisfit=mean(abs(misfit));
[amod,bestmisfit,iter]=BroydenFunc(Broydeniter,dbgain,dCgain,bestmisfit,T1,t,m,w,tetha,...
    dt,target1,b,u,t_peak,C,misfit,errortol,avgtol,amod,iter,wavmag,wavtot,CoffDiag);

%Step 6, Saving best result
[abest,apeak,misfit,bestmisfitall]=BestAcc(amod,abest,t,m,tetha,dt,Tall,wall,targetall,bestmisfitall);

%Step 7, Termination condition
maxmisfit=abs(max((abs(apeak)-targetall')./targetall'));
if mean(abs(misfit))<=avgtol||maxmisfit<errortol; break; end;

%PERIOD SUBSET 2
if T2==0; 
else
w=2*pi./T2; %angular frequency

%Step 1, calculate initial response
[u,t_peak,t_index,apeak,misfit,target]=IniResponse(amod,t,m,tetha,dt,T2,w,target2);

%Step 2, Least square scalling
[amod]=AccScale(apeak,target,amod);

%Step 3: Scale the acceleration time series based on the least square fit
[u,t_peak,t_index,apeak,misfit]=IniResponse(amod,t,m,tetha,dt,T2,w,target2);

%Step 4, calculate initial b, initial C and initial misfit
[b,C,misfit,wavtot]=InitialC2(T2,w,m,tetha,dt,t_peak,t_index,t,Cgain,target2,misfit,amod,CoffDiag,wavmag);

%Step 5, Broyden Loop
bestmisfit=mean(abs(misfit));
[amod,bestmisfit,iter]=BroydenFunc(Broydeniter,dbgain,dCgain,bestmisfit,T2,t,m,w,tetha,...
    dt,target2,b,u,t_peak,C,misfit,errortol,avgtol,amod,iter,wavmag,wavtot,CoffDiag);

%Step 6, Saving best result
[abest,apeak,misfit,bestmisfitall]=BestAcc(amod,abest,t,m,tetha,dt,Tall,wall,targetall,bestmisfitall);

%Step 7, Termination condition
maxmisfit=abs(max((abs(apeak)-targetall')./targetall'));
if mean(abs(misfit))<=avgtol||maxmisfit<errortol; break; end;
end

%PERIOD SUBSET 3
if T3==0; 
else
w=2*pi./T3; %angular frequency

%Step 1, calculate initial response
[u,t_peak,t_index,apeak,misfit,target]=IniResponse(amod,t,m,tetha,dt,T3,w,target3);

%Step 2, Least square scalling
[amod]=AccScale(apeak,target,amod);

%Step 3: Scale the acceleration time series based on the least square fit
[u,t_peak,t_index,apeak,misfit]=IniResponse(amod,t,m,tetha,dt,T3,w,target3);

%Step 4, calculate initial b, initial C and initial misfit
[b,C,misfit,wavtot]=InitialC2(T3,w,m,tetha,dt,t_peak,t_index,t,Cgain,target3,misfit,amod,CoffDiag,wavmag);

%Step 5, Broyden Loop
bestmisfit=mean(abs(misfit));
[amod,bestmisfit,iter]=BroydenFunc(Broydeniter,dbgain,dCgain,bestmisfit,T3,t,m,w,tetha,...
    dt,target3,b,u,t_peak,C,misfit,errortol,avgtol,amod,iter,wavmag,wavtot,CoffDiag);

%Step 6, Saving best result
[abest,apeak,misfit,bestmisfitall]=BestAcc(amod,abest,t,m,tetha,dt,Tall,wall,targetall,bestmisfitall);

%Step 7, Termination condition
maxmisfit=abs(max((abs(apeak)-targetall')./targetall'));
if mean(abs(misfit))<=avgtol||maxmisfit<errortol; break; end;
end

%PERIOD SUBSET 4
if T4==0; 
else
w=2*pi./T4; %angular frequency

%Step 1, calculate initial response
[u,t_peak,t_index,apeak,misfit,target]=IniResponse(amod,t,m,tetha,dt,T4,w,target4);

%Step 2, Least square scalling
[amod]=AccScale(apeak,target,amod);

%Step 3: Scale the acceleration time series based on the least square fit
[u,t_peak,t_index,apeak,misfit]=IniResponse(amod,t,m,tetha,dt,T4,w,target4);

%Step 4, calculate initial b, initial C and initial misfit
[b,C,misfit,wavtot]=InitialC2(T4,w,m,tetha,dt,t_peak,t_index,t,Cgain,target4,misfit,amod,CoffDiag,wavmag);

%Step 5, Broyden Loop
bestmisfit=mean(abs(misfit));
[amod,bestmisfit,iter]=BroydenFunc(Broydeniter,dbgain,dCgain,bestmisfit,T4,t,m,w,tetha,...
    dt,target4,b,u,t_peak,C,misfit,errortol,avgtol,amod,iter,wavmag,wavtot,CoffDiag);

%Step 6, Saving best result
[abest,apeak,misfit,bestmisfitall]=BestAcc(amod,abest,t,m,tetha,dt,Tall,wall,targetall,bestmisfitall);

%Step 7, Termination condition
maxmisfit=abs(max((abs(apeak)-targetall')./targetall'));
if mean(abs(misfit))<=avgtol||maxmisfit<errortol; break; end;
end






end %of outer loop

%Form the response spectra
[u,t_peak,t_index,apeak]=IniResponse(arec,t,m,tetha,dt,Tall,wall,targetall);
Rrec=abs(apeak);
[u,t_peak,t_index,apeak]=IniResponse(abest,t,m,tetha,dt,Tall,wall,targetall);
R=abs(apeak); targetall=abs(targetall);

%Calculate velocity and displacement time series
tae = tf(1,[1 0]);
tae = ss(tae);
voriginal= lsim(tae,arec*9.8103,t,0);   %Original velocity
doriginal=lsim(tae,voriginal,t,0);      %Original displacement
vmod= lsim(tae,abest*9.8103,t,0);       %Modified velocity
dmod=lsim(tae,vmod,t,0);                %Modified displacement
IAori=lsim(tae,(arec*9.8103).^2,t,0)*pi/2/9.8103;
IAproposed=lsim(tae,(abest.*9.8103).^2,t,0)*pi/2/9.8103;

%Ploting
figure(1) %Response Spectrum
plot(Tall,Rrec,'b',Tall,R,'r',Tall,targetall,'k','linewidth',2)
title('Response Spectra');
ylabel('Pseudo-Acceleration (g)');
xlabel('Period (sec)');
axes_handle=get(gcf,'CurrentAxes');
set(axes_handle,'FontSize',10,'FontName','Arial');
set(gca,'LineStyleOrder','-')
set(gcf,'Color',[1,1,1])
legend('Original Ground Motion','Modified Ground Motion','Target')
grid on
 
 
figure(2) %Time series
subplot(2,1,1)
plot(t,arec,'b')
title('(a) Original Acceleration Time Series');
ylabel('Acceleration (g)');
xlabel('Time (sec)');
axes_handle=get(gcf,'CurrentAxes');
set(axes_handle,'FontSize',10,'FontName','Arial');
set(gca,'LineStyleOrder','-')
set(gcf,'Color',[1,1,1])
grid on
 
subplot(2,1,2)
accnew=abest;
tnew=t;
plot(t,abest,'r')
title('(b) Modified Acceleration Time Series');
ylabel('Acceleration (g)');
xlabel('Time (sec)');
axes_handle=get(gcf,'CurrentAxes');
set(axes_handle,'FontSize',10,'FontName','Arial');
set(gca,'LineStyleOrder','-')
set(gcf,'Color',[1,1,1])
grid on
 
% subplot(3,2,3)
% plot(t,voriginal,'b')
% title('(c) Original Velocity Time Series');
% ylabel('Velocivty (m/s)');
% xlabel('Time (sec)');
% axes_handle=get(gcf,'CurrentAxes');
% set(axes_handle,'FontSize',10,'FontName','Arial');
% set(gca,'LineStyleOrder','-')
% set(gcf,'Color',[1,1,1])
% grid on
%  
% subplot(3,2,4)
% plot(t,vmod,'r')
% title('(d) Modified Velocity Time Series');
% ylabel('Velocivty (m/s)');
% xlabel('Time (sec)');
% axes_handle=get(gcf,'CurrentAxes');
% set(axes_handle,'FontSize',10,'FontName','Arial');
% set(gca,'LineStyleOrder','-')
% set(gcf,'Color',[1,1,1])
% grid on
%  
% subplot(3,2,5)
% plot(t,doriginal,'b')
% title('(e) Original Displacement Time Series');
% ylabel('Displacement (m)');
% xlabel('Time (sec)');
% axes_handle=get(gcf,'CurrentAxes');
% set(axes_handle,'FontSize',10,'FontName','Arial');
% set(gca,'LineStyleOrder','-')
% set(gcf,'Color',[1,1,1])
% grid on
%  
% subplot(3,2,6)
% plot(t,dmod,'r')
% title('(f) Modified Displacement Time Series');
% ylabel('Displacement (m)');
% xlabel('Time (sec)');
% axes_handle=get(gcf,'CurrentAxes');
% set(axes_handle,'FontSize',10,'FontName','Arial');
% set(gca,'LineStyleOrder','-')
% set(gcf,'Color',[1,1,1])
% grid on


% figure(3)
% plot(t,IAori,'k',t,IAproposed,'b')
% title('Arias Intensity');
% ylabel('Arias Intensity (m/s)');
% xlabel('t (sec)');
% axes_handle=get(gcf,'CurrentAxes');
% set(axes_handle,'FontSize',10,'FontName','Arial');
% set(gca,'LineStyleOrder','-')
% set(gcf,'Color',[1,1,1])
% legend('Original Ground Motion','Modified Ground Motion','location','SE')
% grid on
% 
% figure(4)
% plot(t,IAori./IAori(end),'k',t,IAproposed./IAproposed(end),'b')
% title('Arias Intensity');
% ylabel('Arias Intensity (Arias/Total Arias)');
% xlabel('t (sec)');
% axes_handle=get(gcf,'CurrentAxes');
% set(axes_handle,'FontSize',10,'FontName','Arial');
% set(gca,'LineStyleOrder','-')
% set(gcf,'Color',[1,1,1])
% legend('Original Ground Motion','Modified Ground Motion','location','SE')
% grid on
% 
