function  [b,Cini,misfitini,wavtot,t_peak]=InitialC2(T,w,m,tetha,increment,t_peak,t_index,t,Cgain,target,misfit,amod,CoffDiag,wavmag)
%This routine calculate the C matrix and initial b vector

%initial set
b=ones(length(T));
tw_index=zeros(1,length(T));
tw_peak=zeros(1,length(T));
dTj=zeros(1,length(T));  
wav=zeros(length(T),length(t));
gamma=0.5;
betha=0.25;

%--------------------------------------------------------------------------
%INITIAL b CALCULATION-----------------------------------------------------
%--------------------------------------------------------------------------
for i=1:length(T) %structure
uw=zeros(1,length(t));
uwdot=0;

    wj=w(i)*sqrt(1-tetha^2);  
    tshift=t_peak(i);
    tj=t-tshift+dTj(i);   
    gf=1.178.*(1/T(i))^-0.93;
    
    wav(i,:)=wavmag.*cos(wj.*tj).*exp(-(tj./gf).^2).*m.*b(i);
     
    ks=m.*w(i)^2;
    c=2*m*w(i)*tetha;
    uwddot=(wav(i,1)-c.*uwdot-ks.*uw(1))/m;

    a1=1/(betha*increment^2).*m+gamma/(betha*increment).*c;
    a2=1/(betha*increment).*m+(gamma/betha-1).*c;
    a3=(1/2/betha-1).*m+increment*(gamma/2/betha-1).*c;
    kh=ks+a1;

    for j=1:length(t)-1
        ph=wav(i,j+1)+a1*uw(j)+a2*uwdot+a3*uwddot;
        uw(j+1)=ph/kh;
        uwdoti=uwdot;
        uwdot=gamma/(betha*increment).*(uw(j+1)-uw(j))+(1-gamma/betha).*uwdot+increment*(1-gamma/2/betha)*uwddot;
        uwddot=1/(betha*increment^2).*(uw(j+1)-uw(j))-1/(betha*increment).*uwdoti-(1/2/betha-1)*uwddot;
    end
    
        % Find peak time of each wavelet   
        abslt=abs(uw).*w(i)^2;
        R=max(abslt);
        tw_index(i)=find(abslt==R);
        tw_peak(i)=(tw_index(i)-1).*increment; 
  
        dTj(i)=tw_peak(i)-t_peak(i);
        tj=t-tshift+dTj(i);
        wav(i,:)=wavmag.*cos(wj.*tj).*exp(-(tj./gf).^2).*m.*b(i);
end

tw_peak=zeros(length(T),length(T));
misfitini=zeros(1,length(T));
C=ones(length(T),length(T));
wavtot=zeros(1,length(t));

for i=1:length(T) %structure

for k=1:length(T) %wavelets
uwdot=0;
 
    gamma=0.5;
    betha=0.25;
    ks=m.*w(i)^2;
    c=2*m*w(i)*tetha;
    uwddot=(wav(k,1)-c.*uwdot-ks.*uw(1))/m;

    a1=1/(betha*increment^2).*m+gamma/(betha*increment).*c;
    a2=1/(betha*increment).*m+(gamma/betha-1).*c;
    a3=(1/2/betha-1).*m+increment*(gamma/2/betha-1).*c;
    kh=ks+a1;
    
        for j=1:length(t)-1
        ph=wav(k,j+1)+a1*uw(j)+a2*uwdot+a3*uwddot;
        uw(j+1)=ph/kh;
        uwdoti=uwdot;
        uwdot=gamma/(betha*increment).*(uw(j+1)-uw(j))+(1-gamma/betha).*uwdot+increment*(1-gamma/2/betha)*uwddot;
        uwddot=1/(betha*increment^2).*(uw(j+1)-uw(j))-1/(betha*increment).*uwdoti-(1/2/betha-1)*uwddot;
        %uddot(i,j+1)=1/(betha*increment^2).*(u(i,j+1)-u(i,j))-1/(betha*increment).*udot(i,j)-(1/2/betha-1)*uddot(i,j);
        end
        
        % Find peak time of each wavelet
        absltw=abs(uw(:)).*w(i)^2;
        Rw=max(absltw);
        tw_index(k)=find(absltw==Rw);
        tw_peak(k,i)=(tw_index(k)-1).*increment; 
        C(i,k)=uw(t_index(i)).*w(i)^2;        
end 
end
% tw_peak=diag(tw_peak);
Cini=Cgain.*C;
% Reduced off-diagonal C matrix
C(setdiff(1:numel(C),1:length(C)+1:numel(C)))=CoffDiag.*C(setdiff(1:numel(C),1:length(C)+1:numel(C))); %suggested value 0.7

% b matrix
b=C\misfit';

atot=amod;
for i=1:length(T)
atot=atot+b(i).*wav(i,:)';
wavtot=wavtot+b(i).*wav(i,:);
end  
clear tw_index;

for i=1:length(T) %structure
uw=zeros(1,length(t));
uwdot=0;
    ks=m.*w(i)^2;
    c=2*m*w(i)*tetha;
    uwddot=(atot(1)-c.*uwdot-ks.*uw(1))/m;

    a1=1/(betha*increment^2).*m+gamma/(betha*increment).*c;
    a2=1/(betha*increment).*m+(gamma/betha-1).*c;
    a3=(1/2/betha-1).*m+increment*(gamma/2/betha-1).*c;
    kh=ks+a1;

    for j=1:length(t)-1
        ph=atot(j+1)+a1*uw(j)+a2*uwdot+a3*uwddot;
        uw(j+1)=ph/kh;
        uwdoti=uwdot;
        uwdot=gamma/(betha*increment).*(uw(j+1)-uw(j))+(1-gamma/betha).*uwdot+increment*(1-gamma/2/betha)*uwddot;
        uwddot=1/(betha*increment^2).*(uw(j+1)-uw(j))-1/(betha*increment).*uwdoti-(1/2/betha-1)*uwddot;
    end
    
        % Find peak time of each wavelet   
        abslt=abs(uw).*w(i)^2;
        R=max(abslt);
        tw_index= abslt==R;
        apeaknew=uw(tw_index)*w(i)^2;
  

if apeaknew<0; target(i)=-abs(target(i)); end   
misfitini(i)=target(i)-apeaknew; 
end
