function [T1,T2,T3,T4,target1,target2,target3,target4]=PeriodSubset(Tall,T1range,T2range,T3range,T4range,targetall)
%This routine forms the period subset(s)

%Period Subset1
[c indexmin]=min(abs(Tall-min(T1range)));
[c indexmax]=min(abs(Tall-max(T1range)));
T1=Tall(indexmin:indexmax);
target1=targetall(indexmin:indexmax);

%Period Subset2
if T2range==0; T2=0;target2=0;
else
[c indexmin]=min(abs(Tall-min(T2range)));
[c indexmax]=min(abs(Tall-max(T2range)));
T2=Tall(indexmin:indexmax);
target2=targetall(indexmin:indexmax);
end

%Period Subset3
if T3range==0; T3=0;target3=0;
else
[c indexmin]=min(abs(Tall-min(T3range)));
[c indexmax]=min(abs(Tall-max(T3range)));
T3=Tall(indexmin:indexmax);
target3=targetall(indexmin:indexmax);
end

%Period Subset4
if T4range==0; T4=0;target4=0;
else
[c indexmin]=min(abs(Tall-min(T4range)));
[c indexmax]=min(abs(Tall-max(T4range)));
T4=Tall(indexmin:indexmax);
target4=targetall(indexmin:indexmax);
end