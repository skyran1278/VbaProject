clc; clear; close all;

structure_period = 1.01;
ss_s1 = [0.5, 0.3, 0.7, 0.4];

% all(42) time history
% filenames = [
%   "RSN68_SFERN_PEL090", "RSN68_SFERN_PEL180", "RSN125_FRIULI.A_A-TMZ000", "RSN125_FRIULI.A_A-TMZ270", "RSN169_IMPVALL.H_H-DLT262", "RSN169_IMPVALL.H_H-DLT352", "RSN174_IMPVALL.H_H-E11140", "RSN174_IMPVALL.H_H-E11230", "RSN721_SUPER.B_B-ICC000", "RSN721_SUPER.B_B-ICC090", "RSN725_SUPER.B_B-POE270", "RSN725_SUPER.B_B-POE360", "RSN752_LOMAP_CAP000", "RSN752_LOMAP_CAP090", "RSN767_LOMAP_G03000", "RSN767_LOMAP_G03090", "RSN848_LANDERS_CLW-LN", "RSN848_LANDERS_CLW-TR", "RSN900_LANDERS_YER270", "RSN900_LANDERS_YER360", "RSN953_NORTHR_MUL009", "RSN953_NORTHR_MUL279", "RSN960_NORTHR_LOS000", "RSN960_NORTHR_LOS270", "RSN1111_KOBE_NIS000", "RSN1111_KOBE_NIS090", "RSN1116_KOBE_SHI000", "RSN1116_KOBE_SHI090", "RSN1148_KOCAELI_ARE000", "RSN1148_KOCAELI_ARE090", "RSN1158_KOCAELI_DZC180", "RSN1158_KOCAELI_DZC270", "RSN1244_CHICHI_CHY101-E", "RSN1244_CHICHI_CHY101-N", "RSN1485_CHICHI_TCU045-E", "RSN1485_CHICHI_TCU045-N", "RSN1602_DUZCE_BOL000", "RSN1602_DUZCE_BOL090", "RSN1633_MANJIL_ABBAR--L", "RSN1633_MANJIL_ABBAR--T", "RSN1787_HECTOR_HEC000", "RSN1787_HECTOR_HEC090"
% ];

% NM = [
%   1.910, 2.449, 1.815, 1.359, 1.576, 1.257, 1.152, 0.930, 0.863, 0.993, 1.008, 1.430, 1.091, 1.401, 1.143, 0.913, 1.502, 0.955, 0.811, 1.426, 0.700, 0.622, 0.935, 1.009, 0.886, 1.084, 1.324, 1.902, 2.973, 1.035, 0.705, 0.745, 0.638, 0.380, 0.828, 0.894, 0.742, 0.630, 0.977, 0.820, 1.595, 0.926
% ];

% % 21 time history
filenames = [
  "RSN68_SFERN_PEL090", "RSN125_FRIULI.A_A-TMZ270", "RSN169_IMPVALL.H_H-DLT262", "RSN174_IMPVALL.H_H-E11230", "RSN721_SUPER.B_B-ICC090", "RSN725_SUPER.B_B-POE360", "RSN752_LOMAP_CAP000", "RSN767_LOMAP_G03090", "RSN848_LANDERS_CLW-TR", "RSN900_LANDERS_YER270", "RSN953_NORTHR_MUL279", "RSN960_NORTHR_LOS000", "RSN1111_KOBE_NIS000", "RSN1116_KOBE_SHI000", "RSN1148_KOCAELI_ARE090", "RSN1158_KOCAELI_DZC180", "RSN1244_CHICHI_CHY101-N", "RSN1485_CHICHI_TCU045-E", "RSN1602_DUZCE_BOL090", "RSN1633_MANJIL_ABBAR--T", "RSN1787_HECTOR_HEC090"
];

NM = [
  2.054, 1.462, 1.695, 1.000, 1.067, 1.537, 1.173, 0.982, 1.027, 0.873, 0.669, 1.005, 0.953, 1.424, 1.113, 0.758, 0.408, 0.891, 0.677, 0.882, 0.996
];

% pick 10 time history
% filenames = [
%   "RSN125_FRIULI.A_A-TMZ000", "RSN767_LOMAP_G03000", "RSN1148_KOCAELI_ARE000", "RSN1602_DUZCE_BOL000", "RSN1111_KOBE_NIS090", "RSN1633_MANJIL_ABBAR--L", "RSN725_SUPER.B_B-POE270", "RSN68_SFERN_PEL180", "RSN960_NORTHR_LOS270", "RSN1485_CHICHI_TCU045-N"
% ];

% NM = [
%   1.737, 1.093, 2.845, 0.710, 1.037, 0.935, 0.964, 2.343, 0.965, 0.856
% ];

tn = 0.001 : 0.01 : 5;
tn_length = length(tn);
acceleration = zeros(length(filenames), tn_length);

for filename = 1 : length(filenames)
    [ag, time_interval, NPTS, errCode] = parseAT2('../PEERNGARecords_Unscaled/' + filenames(filename) + '.AT2');

    period = 0 : time_interval : (NPTS - 1) * time_interval;

    for index = 1 : tn_length

        [~, ~, a_array] = newmark_beta(ag, time_interval, 0.05, tn(index), 'average');

        acceleration(filename, index) = max(abs(a_array));

    end

    acceleration(filename, :) = acceleration(filename, :) * NM(filename);

end

median_acceleration = median(acceleration);

% 0.5 0.3 0.7 0.4
[sad, ~] = design_spectrum(ss_s1(1), ss_s1(2), 0.05, tn);
[sam, ~] = design_spectrum(ss_s1(3), ss_s1(4), 0.05, tn);

figure;
xlabel('T(sec)');
ylabel('Sa(g)');
hold on;

plot(tn, sad);
plot(tn, sam);
plot(tn, median_acceleration);

tol = 0.001;

period_index = abs(tn - structure_period) < tol;

sa = median_acceleration(period_index);

fprintf('Sa: %.3f, SaD Factor: %.3f, SaM Factor: %.3f\n', sa, sad(period_index) / sa, sam(period_index) / sa);
% fprintf('Records: %s, PGA: %.3f, PGA: %.3f\n', max(abs(ag)), acceleration(1));
