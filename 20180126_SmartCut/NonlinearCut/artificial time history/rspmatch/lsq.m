function S=lsq(AccScale,apeak,target)
%This routine calculate the least square fit of the spectral misfit

apeak=AccScale.*apeak;
apeak=abs(apeak');
target=abs(target);

S=sum((apeak-target).^2);


