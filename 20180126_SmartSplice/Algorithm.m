clc; clear; close all;

topAs = [9 4 10] * 8.19;

botAs = [10 5 9] * 5.09;

band = 3.50;

span = 10.00;

DL = 0.337 * band;

LL = 0.2 * band;

fy = 4200 / 1000; % tf

d = (80 - 4 - 1.27) / 100; % m

topMn = topAs * fy * 0.9 * d; % tf-m

botMn = botAs * fy * 0.9 * d; % tf-m

x = 0 : 0.01 : span;
x_length = length(x);

midline = zeros(1, x_length);

DLMn = 4 * (1 / 8 * DL * span ^ 2) * (x / span - 0.5) .^ 2 - (1 / 24 * DL * span ^ 2);
LLMn = 4 * (1 / 8 * LL * span ^ 2) * (x / span - 0.5) .^ 2 - (1 / 24 * LL * span ^ 2);

maxEQ = min(botMn(1), botMn(3));

EQ = -maxEQ + 2 * maxEQ / span * x;
NEQ = -EQ;

settle = -10 + 2 * 10 / span * x;

% DL + EQ
negativeMoment = EQ .* (EQ >= 0) + NEQ .* (NEQ >= 0) + 1.2 * DLMn .* (DLMn >= 0) + 0.5 * LLMn .* (LLMn >= 0);
positiveMoment = EQ .* (EQ <= 0) + NEQ .* (NEQ <= 0) + 1.2 * DLMn .* (DLMn <= 0) + 0.5 * LLMn .* (LLMn <= 0);

greenColor = [26 188 156] / 256;
blueColor = [52 152 219] / 256;
redColor = [233 88 73] / 256;
grayColor = [0.5 0.5 0.5];

figure;
hold on;
plot(x, midline, '-k');
% plot(x, 1.4 * DLMn, 'Color', redColor);
% plot(x, 1.2 * DLMn + 1.6 * LLMn, 'Color', redColor);
% plot(x, 1.2 * DLMn + 0.5 * LLMn, 'Color', blueColor);
plot(x, 1.2 * DLMn + 0.5 * LLMn + settle, 'Color', blueColor);
% plot(x, EQ, 'Color', greenColor);
% legendEQ = plot(x, NEQ, 'Color', greenColor);
% legendMoment = plot(x, positiveMoment, 'Color', redColor);
% plot(x, negativeMoment, 'Color', redColor);
% plot(x, LLMn, 'Color', blueColor);
% plot([0, span / 2, span], topMn, 'o', 'Color', redColor);
% legendAs = plot([0, span / 2, span], -botMn, 'o', 'Color', redColor);
axis([0 span -250 250]);
title('Moment');
xlabel('m');
ylabel('tf-m');
