function [u,t_peak,t_index,apeak,misfit,target]=IniResponse(a,t,m,tetha,increment,Tn,wn,target)
%This routine calculates the oscilator response using Newmarks Method

%Initial set
t_index=zeros(1,length(Tn));
t_peak=zeros(1,length(Tn));
apeak=zeros(1,length(Tn));
misfit=zeros(1,length(Tn));
u=zeros(length(Tn),length(t));

%Newmark coefficient
gamma=0.5;
betha=0.25;

for i=1:length(Tn)
udot=0;
ks=m.*wn(i)^2;
c=2*m*wn(i)*tetha;
uddot=(a(1)-c.*udot-ks.*u(i,1))/m;

a1=1/(betha*increment^2).*m+gamma/(betha*increment).*c;
a2=1/(betha*increment).*m+(gamma/betha-1).*c;
a3=(1/2/betha-1).*m+increment*(gamma/2/betha-1).*c;
kh=ks+a1;

    for j=1:length(t)-1
        ph=a(j+1)+a1*u(i,j)+a2*udot+a3*uddot;
        u(i,j+1)=ph/kh;
        udoti=udot;
        udot=gamma/(betha*increment).*(u(i,j+1)-u(i,j))+(1-gamma/betha).*udot+increment*(1-gamma/2/betha)*uddot;
        uddot=1/(betha*increment^2).*(u(i,j+1)-u(i,j))-1/(betha*increment).*udoti-(1/2/betha-1)*uddot;
     end
    
abslt=abs(u(i,:))*wn(i)^2;              %Absolute pseudo acceleration response (psa)
R=max(abslt);                           %Max of absolute psa
t_index(i)=find(abslt==R);              %Index of max absolute psa on time series
t_peak(i)=(t_index(i)-1).*increment;    %Time of peak psa
apeak(i)=u(i,t_index(i)).*wn(i)^2;      %True value of max peak
end

%Calculate misfit
for i=1:length(Tn)
if apeak(i)<0; target(i)=-abs(target(i)); end   
misfit(i)=target(i)-apeak(i); 
end

