clc; clear; close all;

SaD = 0.6;
SaM = 0.8;

T = 0.344; % from etabs
T0D = 1.6;
T0M = 1.6;

framing_type = 2;
performance_level = 'LS';

theta = 0.01;
k = cal_k(T)
VD = c1(T, T0D) * c2(T, T0D, framing_type, performance_level) * c3(T, theta) * SaD
VM = c1(T, T0M) * c2(T, T0M, framing_type, performance_level) * c3(T, theta) * SaM


function output = cal_k(T)
    if T <= 0.5
        output = 1;
    elseif T >= 2.5
        output = 2;
    else
        output = (2 - 1) / (2.5 - 0.5) * (T - 0.5);
    end
end
