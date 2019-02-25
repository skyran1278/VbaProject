clc; clear; close all;

SaD = 0.6;
SaM = 0.8;

T = 0.344; % from etabs
T0D = 1.6;
T0M = 1.6;

SDS = 0.6;
SMS = 0.8;
SD1 = SaD * T0D;
SM1 = SaM * T0M;

framing_type = 2;
performance_level = 'LS';
theta = 0.01;

theta = 0.01;
k = cal_k(T)
VD = c1(T, T0D) * c2(T, T0D, framing_type, performance_level) * c3(T, theta) * SaD
VM = c1(T, T0M) * c2(T, T0M, framing_type, performance_level) * c3(T, theta) * SaM

ki = 6.483;
Vy = 133.143;
ke = 6.165;
Te = T * sqrt(ki / ke);

stories_number = 3

W =

RD = SaD / (Vy / W) / c0(stories_number)

c1_nsp(Te, T0D, RD, c1(T, T0D))
