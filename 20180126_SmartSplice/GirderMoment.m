clc; clear; close all;

maxEQ = 30;
positiveDL = 10;
negativeDL = 10;

beamLength = 10;

x = 0 : 0.01 : beamLength;
x_length = length(x);

midline = zeros(1, x_length);

EQ = -maxEQ + 2 * maxEQ / beamLength * x;
NEQ = -EQ;
DL = -positiveDL + 4 * (positiveDL + negativeDL) * (x / beamLength - 0.5) .^ 2;

% DL + EQ
negativeMoment = EQ .* (EQ >= 0) + NEQ .* (NEQ >= 0) + 1.4 * DL .* (DL >= 0);
positiveMoment = EQ .* (EQ <= 0) + NEQ .* (NEQ <= 0) + 1.4 * DL .* (DL <= 0);

topLeftRebar = max(negativeMoment(x <= beamLength / 3));
topRightRebar = max(negativeMoment(x >= beamLength / 3));

% 負彎矩需求曲線
topRebar = [topLeftRebar - topLeftRebar / (beamLength / 2) * x(x <= beamLength / 2), topRightRebar / (beamLength / 2) * x(x > beamLength / 2) - topRightRebar];

% 取正
botLeftRebar = -min(positiveMoment(x <= beamLength / 3));
botMidRebar = -min(positiveMoment(x >= 1 * beamLength / 4 & x <= 3 * beamLength / 4));
botRightRebar = -min(positiveMoment(x >= 2 * beamLength / 3));

botRebarDL = 4 * botMidRebar * (x / beamLength - 0.5) .^ 2 - botMidRebar;

botRebar = min([EQ; NEQ; botRebarDL]);
botRebarOtherMethod = [-botLeftRebar + (botLeftRebar - botMidRebar) / (beamLength / 2) * x(x <= beamLength / 2), -botMidRebar + (botMidRebar - botRightRebar) / (beamLength / 2) * (x(x > beamLength / 2) - (beamLength / 2)) ];

% bot = - botMidRebar * ones(1, x_length);

% figure;
% plot(x, midline, '-k');
% hold on;
% plot(x, EQ, ':k', x, NEQ, ':k');
% plot(x, DL, ':k');
% plot(x, negativeMoment, '--r', x, positiveMoment, '--r');
% % plot(x, botRebarDL, '--b');
% % plot(x, bot, '--g');
% plot(x, topRebar, '-b');
% plot(x, botRebar, '-b');
% plot(x, botRebarOtherMethod, '-g');

% plot(0, topLeftRebar, 'or');
% plot(beamLength / 2, 0, 'or');
% plot(beamLength, topRightRebar, 'or');
% plot(0, -botLeftRebar, 'or');
% plot(beamLength / 2, -botMidRebar, 'or');
% plot(beamLength, -botRightRebar, 'or');

figure;
plot(x, midline, '-k');
hold on;
legendEQ = plot(x, EQ, '-r', x, NEQ, '-r');
legendGravity = plot(x, DL, '-b');
axis([0 beamLength -50 50]);
legend([legendEQ(1), legendGravity], 'EQ', 'Gravity', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

figure;
plot(x, midline, '-k');
hold on;
legendEQ = plot(x, EQ, '-k', x, NEQ, '-k');
legendGravity = plot(x, DL, '-k');
legendMoment = plot(x, negativeMoment, '-b', x, positiveMoment, '-b');
axis([0 beamLength -50 50]);
legend([legendEQ(1), legendGravity, legendMoment(1)], 'EQ', 'Gravity', 'Linear Add', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

figure;
plot(x, midline, '-k');
hold on;
% legendEQ = plot(x, EQ, '-k', x, NEQ, '-k');
% legendGravity = plot(x, DL, '-k');
legendMoment = plot(x, negativeMoment, '-k', x, positiveMoment, '-k');
% % plot(x, botRebarDL, '--b');
% % plot(x, bot, '--g');
% plot(x, topRebar, '-b');
% plot(x, botRebar, '-b');
% plot(x, botRebarOtherMethod, '-g');

legendActural = plot(0, topLeftRebar, 'ob');
plot(beamLength / 2, 0, 'ob');
plot(beamLength, topRightRebar, 'ob');
plot(0, -botLeftRebar, 'ob');
plot(beamLength / 2, -botMidRebar, 'ob');
plot(beamLength, -botRightRebar, 'ob');
axis([0 beamLength -50 50]);
legend([legendActural, legendMoment(1)], 'Rebar', 'Demand', 'Location', 'northeast');
title('Moment');
xlabel('m');
ylabel('tf-m');

% figure;
% plot(x, midline, '-k');
% hold on;
% legendEQ = plot(x, EQ, '-k', x, NEQ, '-k');
% legendGravity = plot(x, DL, '-k');
% legendMoment = plot(x, negativeMoment, '-r', x, positiveMoment, '-r');
% % % plot(x, botRebarDL, '--b');
% % % plot(x, bot, '--g');
% % plot(x, topRebar, '-b');
% % plot(x, botRebar, '-b');
% % plot(x, botRebarOtherMethod, '-g');

% plot(0, topLeftRebar, 'or');
% plot(beamLength / 2, 0, 'or');
% plot(beamLength, topRightRebar, 'or');
% plot(0, -botLeftRebar, 'or');
% plot(beamLength / 2, -botMidRebar, 'or');
% plot(beamLength, -botRightRebar, 'or');
% axis([0 beamLength -50 50]);
% legend([legendEQ(1), legendGravity, legendMoment(1)], 'EQ', 'Gravity', 'Linear Add', 'Location', 'northeast');
% title('Moment');
% xlabel('m');
% ylabel('tf-m');
