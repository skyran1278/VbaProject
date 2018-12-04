% 看看是否如我的預期
clc; clear; close all;

varLeft = 50.9;
varMid = 25.5;
varRight = 45.8;

varGravity = 10.5;

beamLength = 10;

ratio = 1 : -0.01 : 0;
ratioLength = length(ratio);
x = linspace(0, beamLength / 2, ratioLength);

arrGravity = - (1 - ratio .^ 2) * varGravity;
arrEQ = - ratio * varLeft;
arrline = - (ratio * (varLeft - varMid) + varMid)

greenColor = [26 188 156] / 256;
blueColor = [52 152 219] / 256;
redColor = [233 88 73] / 256;
grayColor = [0.5 0.5 0.5];
bgColor = [247 247 247] / 256;

% 1
figure;
hold on;
% plot(x, DL, '-.', 'Color', greenColor, 'LineWidth', 1.75);
plot(x, arrGravity, '--', 'Color', greenColor, 'LineWidth', 1.75);
plot(x, arrEQ, '--', 'Color', blueColor, 'LineWidth', 1.75);
plot(x, arrGravity + arrEQ, '--', 'Color', redColor, 'LineWidth', 1.75);
plot(x, arrline, '-', 'Color', redColor, 'LineWidth', 1.75);
