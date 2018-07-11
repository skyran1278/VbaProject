clc; clear; close all;

beamLength = 1;

x = 0 : 0.01 : beamLength;

x_length = length(x);

midline = zeros(1, x_length);

EQ = -30 + 60 * x / beamLength;

NEQ = 30 - 60 * x / beamLength;

DL = 80 * (x / beamLength - 0.5) .^ 2 - 10;

negativeMoment = EQ .* (EQ >= 0) + NEQ .* (NEQ >= 0) + DL .* (DL >= 0);
positiveMoment = EQ .* (EQ <= 0) + NEQ .* (NEQ <= 0) + DL .* (DL <= 0);

topLeftRebar = max(negativeMoment(x <= beamLength / 3));
topRightRebar = max(negativeMoment(x >= beamLength / 3));

topRebar = [topLeftRebar - topLeftRebar / (beamLength / 2) * x(x <= beamLength / 2), topRightRebar / (beamLength / 2) * x(x > beamLength / 2) - topRightRebar];

botLeftRebar = min(positiveMoment(x <= beamLength / 3));
% botMidRebar = min(positiveMoment(1 * beamLength / 4 <= x <= 3 * beamLength / 4));
botMidRebar = min(positiveMoment(x >= 1 * beamLength / 4 & x <= 3 * beamLength / 4));
botRightRebar = min(positiveMoment(x >= 2 * beamLength / 3));

botRebarDL = -4 * botMidRebar * (x / beamLength - 0.5) .^ 2 + botMidRebar;
% botRebarEQ = [botLeftRebar - botLeftRebar / (beamLength / 2) * x, max(positiveMoment(x > beamLength / 2)) / (beamLength / 2) * x];

figure;
plot(x, midline, '-k');
hold on;
plot(x, EQ, '-k', x, NEQ, '-k');
plot(x, DL, '-k');
plot(x, negativeMoment, '-r', x, positiveMoment, '-r');
plot(x, topRebar, '--b', x, botRebarDL, '--b');
plot(0, topLeftRebar, 'ob');
plot(beamLength / 2, 0, 'ob');
plot(beamLength, topRightRebar, 'ob');
plot(0, botLeftRebar, 'ob');
plot(beamLength / 2, botMidRebar, 'ob');
plot(beamLength, botRightRebar, 'ob');
title('');
xlabel('');
ylabel('');
