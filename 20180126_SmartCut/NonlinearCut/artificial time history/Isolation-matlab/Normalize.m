% write in too long ago, very hard to read 2018/12/17
clc; clear; close all;

input = 'HWA015_PGA.txt';
output = 'HWA015_PGA_Normalize v1.txt';
PGA = dlmread(input, '', 11, 0);

max_PGA = max(max(abs(PGA(:,[2 3 4]))),[],2);

PGA(:,2) = PGA(:,2) / max_PGA;
PGA(:,3) = PGA(:,3) / max_PGA;
PGA(:,4) = PGA(:,4) / max_PGA;

% PGA = reshape(PGA(:, 4), 1, 15000);

if max(abs(PGA(:,2))) == 1
    dlmwrite(output,PGA(:,[1 2]));

elseif max(abs(PGA(:,3))) == 1
    dlmwrite(output,PGA(:,[1 3]));

elseif max(abs(PGA(:,4))) == 1
    dlmwrite(output,PGA(:,[1 4]));

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
