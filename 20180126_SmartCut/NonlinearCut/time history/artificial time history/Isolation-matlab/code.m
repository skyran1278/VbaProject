function Value=code(Ss,S1,F)


T0=S1/Ss;


LF=length(F);
for nf=1:LF
T=1/F(nf);
if T < 0.2*T0
Ca=(0.4+3*T/T0)*Ss;
elseif T<T0
Ca=Ss;
else
Ca=S1/T;
end


Value(nf,:)=[F(nf) Ca];
end