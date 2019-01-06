% write in too long ago, very hard to read 2018/12/17
clc; clear; close all;

input = 'elcentro_EW.txt';
output = 'elcentro_max.txt';

PGA = dlmread(input, '', 11, 0);

NSEW = [3 4];

% only pick x y max
[max_PGA, argmax] = max(max(abs(PGA(:, NSEW))), [], 2);

% gal scale to g
% PGA(:, [2 3 4]) = PGA(:, [2 3 4]) / 980.665;
PGA(:, [2 3 4]) = PGA(:, [2 3 4]) / 100;

% write to file
% .7f
dlmwrite(output, PGA(:, [1 NSEW(argmax)]), 'delimiter', ' ', 'precision', '%.7f');

% normalize
% PGA(:,2) = PGA(:,2) / max_PGA;
% PGA(:,3) = PGA(:,3) / max_PGA;
% PGA(:,4) = PGA(:,4) / max_PGA;

% % PGA = reshape(PGA(:, 4), 1, 15000);

% if max(abs(PGA(:,2))) == 1
%     dlmwrite(output,PGA(:,[1 2]));

% elseif max(abs(PGA(:,3))) == 1
%     dlmwrite(output,PGA(:,[1 3]));

% elseif max(abs(PGA(:,4))) == 1
%     dlmwrite(output,PGA(:,[1 4]));

% end

% plot all time history
figure;
% hold on;
plot(PGA(:,1), PGA(:,[2 3 4]));
% plot(PGA(:,1), PGA(:,3));
% plot(PGA(:,1), PGA(:,4));
legend('U(+)','N(+)','E(+)');
xlabel('Period (sec)');
ylabel('Acceleration (g)');
