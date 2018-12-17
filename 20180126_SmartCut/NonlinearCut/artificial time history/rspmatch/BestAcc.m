function [abest,apeak,misfit,bestmisfitall]=BestAcc(amod,abest,t,m,tetha,dt,Tall,wall,targetall,bestmisfitall)

%This routine saves the best acceleration that gives the smallest average misfit during the Outer-loop

[u,t_peak,t_index,apeak,misfit]=IniResponse(amod,t,m,tetha,dt,Tall,wall,targetall);
if bestmisfitall>=mean(abs(misfit)), bestmisfitall=mean(abs(misfit)); abest=amod; end;