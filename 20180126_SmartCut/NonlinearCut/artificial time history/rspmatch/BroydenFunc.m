function [amod,bestmisfit,iter,b,C,deltab,misfit]=BroydenFunc(Broydeniter,dbgain,dCgain,bestmisfit,T,t,m,w,tetha,increment,...
    target,b,u,t_peak,C,misfit,tol,normtol,amod,iter,wavmag,wavtot,CoffDiag)

%This routine solves non-linear system using Broyden's Updating
%Oscillator response is calculated using Newmark's Method

a=amod;
gamma=0.5;
betha=0.25;
amod=a+wavtot';
b=b';
P=-1;
for q=1:Broydeniter %Broyden loop 
unew=zeros(length(T),length(t));
upeak_index=zeros(1,length(T));
error=zeros(1,length(T));
unew_peak=zeros(1,length(T));
unewpeak_time=zeros(1,length(T));
newmisfit=zeros(length(T),1);
deltamisfit=zeros(length(T),1);
wavtot=zeros(1,length(t)); %resetting total adjustment wavelets

previousmeanmisfit=mean(abs(misfit));
Coff=P*C;
Coff(setdiff(1:numel(Coff),1:length(Coff)+1:numel(Coff)))=CoffDiag.*Coff(setdiff(1:numel(Coff),1:length(Coff)+1:numel(Coff))); %suggested value 0.7
deltab=dbgain*(Coff\misfit');
ns=deltab'*deltab; 
b=b+deltab';

for j=1:length(T)    %wavelet
uw=zeros(1,length(t));
uwdot=0;

    wj=w(j)*sqrt(1-tetha^2);  
    tshift=t_peak(j);%min(t_peak(j),3.9223*(1/Tn(j))^-0.845);
    tj=t-tshift;   
    gf=1.178.*(1/T(j))^-0.93;
    
    wav=wavmag.*b(j).*cos(wj.*tj).*exp(-(tj./gf).^2).*m;
   
    ks=m.*w(j)^2;
    c=2*m*w(j)*tetha;
    uwddot=(wav(1)-c.*uwdot-ks.*uw(1))/m;
    
    a1=1/(betha*increment^2).*m+gamma/(betha*increment).*c;
    a2=1/(betha*increment).*m+(gamma/betha-1).*c;
    a3=(1/2/betha-1).*m+increment*(gamma/2/betha-1).*c;
    kh=ks+a1;
    
           for k=1:length(t)-1
            ph=wav(k+1)+a1*uw(k)+a2*uwdot+a3*uwddot;
            uw(k+1)=ph/kh;
            uwdoti=uwdot;
            uwdot=gamma/(betha*increment).*(uw(k+1)-uw(k))+(1-gamma/betha).*uwdot+increment*(1-gamma/2/betha)*uwddot;
            uwddot=1/(betha*increment^2).*(uw(k+1)-uw(k))-1/(betha*increment).*uwdoti-(1/2/betha-1)*uwddot;
           end
                      
    %Find peak time of wavelet response
    abslt=abs(uw);                      %Absolute wavelet pseudo acceleration response (psa)
    R=max(abslt);                       %Max of absolute psa
    tw_index=find(abslt==R);            %Index of max absolute psa on time series
    tw_peak=(tw_index-1).*increment;    %true value of total max peak     
    dTj=tw_peak-t_peak(j);              %time shifting

    %Shifted wavelet
    tj=t-tshift+dTj; 
    wav=wavmag.*b(j).*cos(wj.*tj).*exp(-(tj./gf).^2).*m;
    wavtot=wavtot+wav';                 %Total wavelet
    
end %end of j (wavelet)

for i=1:length(T)    %structure
uw=zeros(1,length(t));
uwdot=0;

    ks=m.*w(i)^2;
    c=2*m*w(i)*tetha;
    uwddot=(wavtot(1)-c.*uwdot-ks.*uw(1))/m;

    a1=1/(betha*increment^2).*m+gamma/(betha*increment).*c;
    a2=1/(betha*increment).*m+(gamma/betha-1).*c;
    a3=(1/2/betha-1).*m+increment*(gamma/2/betha-1).*c;
    kh=ks+a1;
    
           for j=1:length(t)-1
            ph=wavtot(j+1)+a1*uw(j)+a2*uwdot+a3*uwddot;
            uw(j+1)=ph/kh;
            uwdoti=uwdot;
            uwdot=gamma/(betha*increment).*(uw(j+1)-uw(j))+(1-gamma/betha).*uwdot+increment*(1-gamma/2/betha)*uwddot;
            uwddot=1/(betha*increment^2).*(uw(j+1)-uw(j))-1/(betha*increment).*uwdoti-(1/2/betha-1)*uwddot;
           end
    
    %Total response
    unew(i,:)=u(i,:)+uw;
    
    %Calculate new misfit and deltamisfit
    abslt=abs(unew(i,:));
    R=max(abslt);
    upeak_index(i)=find(abslt==R);
    unew_peak(i)=unew(i,upeak_index(i)); 
    unewpeak_time(i)=(upeak_index(i)-1)*increment;
    
    if unew_peak(i)>=0, target(i)=abs(target(i));
    else target(i)=-abs(target(i));
    end
  
    newmisfit(i)=target(i)-unew_peak(i)*w(i)^2; %New misfit
    deltamisfit(i)=newmisfit(i)-misfit(i);      %delta misfit
    misfit(i)=newmisfit(i);                     %Save new misfit as initial misfit for next iteration
    error(i)=max(abs(target(i))-abs(unew_peak*w(i)^2))/abs(target(i));
end % i loop (structure)

%Broyden Updating
C=C+dCgain*((deltamisfit-C*deltab)*deltab')/ns;

meanmisfit=mean(abs(misfit));
maxerror=max(error);

fprintf('Iter %2.0f. Mean Misfit is %7.4f \n', iter, meanmisfit)

%Save the best modified acceleration for each subperiod solved
if bestmisfit>=meanmisfit, bestmisfit=meanmisfit; amod=a+wavtot'; 
end;
iter=iter+1;

%Termination condition
if meanmisfit>previousmeanmisfit;P=-P; end; % I added this to aid solution convergence. In case that the misfit is diverging, the Broyden will change it direction
if meanmisfit>1e3;break; end;
if meanmisfit<=normtol||maxerror<tol; break, end;
end %end of broyden
