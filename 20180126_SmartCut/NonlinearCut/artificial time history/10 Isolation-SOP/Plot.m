clc;clear all;close all;
%PGA = dlmread('GIAN331.txt','',11,0) ;
%PGA(:,2) = PGA(:,2)/max(PGA(:,4));
%PGA(:,3) = PGA(:,3)/max(PGA(:,4));
%PGA(:,4) = PGA(:,4)/max(PGA(:,4));
%PGA = reshape(PGA(:,4),1,15000);
%plot(PGA(:,1),PGA(:,2));
%hold on;
%plot(PGA(:,1),PGA(:,3));
%plot(PGA(:,1),PGA(:,4));
%xlabel('Time(s)');
%ylabel('gal. DCoffset(corr)');
%title('HW7');
%grid on;
%legend('U','N','E','location','SouthEast');
%dlmwrite('GIAN331_Normal.txt',PGA(:,4),'delimiter','\n');
%type PGAE.txt; %print
%saveas(figure(1),'HW3.1.jpg')
Spectrum = dlmread('SPEC.prn','',20,0);
%dlmwrite('PGAE.txt',PGA(:,4),'delimiter','\n');
plot(Spectrum(:,1),Spectrum(:,6));
