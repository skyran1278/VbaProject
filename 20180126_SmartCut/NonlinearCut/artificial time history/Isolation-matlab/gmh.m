clear
clf
clc

%% EQ


data=load('chichi_TAP010 max ag.txt');

t=data(:,1);
dt=t(2)-t(1);
Lt=length(t);

plot(t,data(:,2))
%axis([0,30,-10,10])
[aa,bb]=ginput(1);
Nv=fix(aa/dt);
vv=1:Nv;


%% Frequency
LN=length(t);
df=1/(LN*dt);
%MF=4;
MF=fix(0.1/df);
Fq=(MF:LN/2)'*df;
LF=length(Fq);
qqq=[Fq zeros(LF,1)];
save Freq.dat qqq -ascii

%% Code

%%%%%%%% zone 4 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

SDS=0.6;
SD1=0.78;
Fa=1.0;
Fv=1.0;
Ss=Fa*SDS;
S1=Fv*SD1;
EPA=0.4*Ss;
Res=code(Ss,S1,Fq);
Code=Res(:,2);


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
x0=data(:,2);
x0=x0/max(abs(x0))*EPA;
x=x0;


%%% Initial Spectrum
qqq=[t x0];
save at.dat qqq -ascii
!spectrum.exe
load spe.dat
F=spe(:,1);
Sa0=spe(:,2);


%%% Do loop
for loop=1:10
Loop=loop

for ii=1:10
x(vv)=0.95*x(vv);
x=x-mean(x);
end

qqq=[t x];
save at.dat qqq -ascii


%% Determine Spectrum
!spectrum.exe

%% Modification Factor
load spe.dat
F=spe(:,1);
Sa=spe(:,2);
Mod=Code./Sa;

%% FFT & Recover
X=fft(x);
XP=angle(X);
XA=abs(X);
XA0=XA(MF+1:LF+MF);
XA1=Mod.*XA0;

XA(MF+1:LF+MF)=XA1;
nn=(LN-MF+1):-1:(LN-LF-MF+2);
XA(nn)=XA1;

XX=XA.*exp(i*XP);
xx=ifft(XX);
x=real(xx);

end  % end of loop

qqq=[t x];
save at.dat qqq -ascii

%% See Results

!spectrum.exe
load spe.dat
F=spe(:,1);
Sa=spe(:,2);

%%subplot(211)
figure(1)
plot(t,x0,t,x)
legend('Initial','Simulated')
%axis([0,90,-1.2,1.2])
xlabel('Time (sec)')
ylabel('Acceleration (g)')

%%subplot(212)
figure(2)
plot(1./F,Sa0,1./F,Sa,1./F,Code,'Linewidth',1.5)
legend('Initial','Simulated','Code')
%axis([0,8,0,1.0])
xlabel('Period (sec)')
ylabel('Sa (g)')
