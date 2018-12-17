function [arec,t]=Zeropad(zeropad,arec,dt)

%This routine perfoms the zero padding on the acceleration time series

pad=zeropad/dt;
apad(1:pad)=0; apad(pad+1:pad+length(arec))=arec; apad(length(apad)+1:length(apad)+pad)=0;
t=(0:length(apad)-1).*dt; t=t';
amod=apad'; abest=amod; arec=apad'; clear apad