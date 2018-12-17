function [acc]=AccScale(apeak,target,acc)
%This routine linearly scales the ground motion

optionOPT=optimset('MaxIter',500,'TolFun',1e-15,'TolX',1e-15);
AccScale=1;
[AccScale]=fminsearch(@(AccScale)lsq(AccScale,apeak,target),AccScale,optionOPT); 
acc=AccScale.*acc;
