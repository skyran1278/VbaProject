function [scaleN,scaleE,noscaleN,noscaleE]=externalforce(scaleG,NofQuake)
% 使用盧學長之33*2筆地震歷時資料
a=dir('*.txt');
% N=length(a);
kk=NofQuake;
[dt]=textread(a(kk).name,'%f',1,'headerlines',12);
[groundTUNE]=textread(a(kk).name,'%f','headerlines',11);
noscaleN=zeros(length(groundTUNE)/4,1);
noscaleE=zeros(length(groundTUNE)/4,1);
for i=1:length(groundTUNE)/4
    noscaleN(i)=groundTUNE(4*(i-1)+3);
    noscaleE(i)=groundTUNE(4*(i-1)+4);
end
noscaleN=noscaleN.*0.01; %cm/s^2 to m/s^2
noscaleE=noscaleE.*0.01; 
pgaN=max(abs(noscaleN(:)));
scaleN=(1/pgaN*scaleG).*noscaleN;
pgaE=max(abs(noscaleE(:)));
scaleE=(1/pgaE*scaleG).*noscaleE;

