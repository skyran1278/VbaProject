% 驗證另一種演算法的正確性
clc; clear; close all;


maxDL = 20;
maxEQ = 67;

beamLength = 10;

x = 0 : 0.01 : beamLength;
xLength = length(x);

midline = zeros(1, xLength);

ratio = 1 : -0.01 : 0;
ratioLength = length(ratio);
xAnother = linspace(0, beamLength / 2, ratioLength);

DL = - maxDL * 2 / 3 + 4 * maxDL * (x / beamLength - 0.5) .^ 2;
DLAnother = - (1 - ratio .^ 2) * maxDL + maxDL / 3;
EQ = ratio * maxEQ;

demand = DLAnother .* (DLAnother >= 0) + EQ;

greenColor = [26 188 156] / 256;
blueColor = [52 152 219] / 256;
redColor = [233 88 73] / 256;
grayColor = [0.5 0.5 0.5];
bgColor = [247 247 247] / 256;

% 1
figure;
hold on;
% plot(x, DL, '-.', 'Color', greenColor, 'LineWidth', 1.75);
plot(xAnother, DLAnother, '--', 'Color', blueColor, 'LineWidth', 1.75);
plot(xAnother, EQ, '--', 'Color', redColor, 'LineWidth', 1.75);
plot(xAnother, demand, '--', 'Color', greenColor, 'LineWidth', 1.75);
