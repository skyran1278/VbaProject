clc;clear all;close all;
PGA = dlmread('PGA.txt','',11,0) ;


PGA(:,2) = PGA(:,2)/max(max(abs(PGA(:,[2 3 4]))),[],2);
PGA(:,3) = PGA(:,3)/max(max(abs(PGA(:,[2 3 4]))),[],2);
PGA(:,4) = PGA(:,4)/max(max(abs(PGA(:,[2 3 4]))),[],2);
%PGA = reshape(PGA(:,4),1,15000);
if max(abs(PGA(:,2))) == 1
    dlmwrite('PGA_Normalize.txt',PGA(:,2),'delimiter','\n');
elseif max(abs(PGA(:,3))) == 1
    dlmwrite('PGA_Normalize.txt',PGA(:,3),'delimiter','\n');
elseif max(abs(PGA(:,4))) == 1
    dlmwrite('PGA_Normalize.txt',PGA(:,4),'delimiter','\n');
end
plot(PGA(:,1),PGA(:,2));
hold on;
plot(PGA(:,1),PGA(:,3));
plot(PGA(:,1),PGA(:,4));
grid on;
legend('U','N','E','location','SouthEast');

%xlabel('Time(s)');
%ylabel('gal. DCoffset(corr)');
%title('HW7');
%type PGAE.txt; %print
%saveas(figure(1),'HW3.1.jpg')
%dlmwrite('PGAE.txt',PGA(:,4),'delimiter','\n');

%Spectrum = dlmread('SPEC.prn','',20,0);
%plot(Spectrum(:,1),Spectrum(:,6));
