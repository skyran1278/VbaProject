clc; clear; close all;

config = Config;

scaled_factor = 0.5 : 0.5 : 10;

scaled_factor_length = length(scaled_factor);

sd_triangle = NaN(1, scaled_factor_length);
sd_uniform = NaN(1, scaled_factor_length);
sd_power = NaN(1, scaled_factor_length);
sa_triangle = NaN(1, scaled_factor_length);
sa_uniform = NaN(1, scaled_factor_length);
sa_power = NaN(1, scaled_factor_length);

for index = 1 : scaled_factor_length
    [sd_triangle(index), sa_triangle(index)] = procedure_b(config, 'triangle', scaled_factor(index));
end

for index = 1 : scaled_factor_length
    [sd_uniform(index), sa_uniform(index)] = procedure_b(config, 'uniform', scaled_factor(index));
end

for index = 1 : scaled_factor_length
    [sd_power(index), sa_power(index)] = procedure_b(config, 'power', scaled_factor(index));
end

green = [26 188 156] / 256;
blue = [52 152 219] / 256;
red = [233 88 73] / 256;
orange = [230 126 34] / 256;
gray = [0.5 0.5 0.5];
background = [247 247 247] / 256;

PF1_phi = 1.278413;
sa = 0.171;

figure;
hold on;
title('ADRS');
xlabel('sd(mm)');
ylabel('sa(g)');
plot(sd_triangle, sa_triangle, 'DisplayName', 'Elastic', 'Color', green, 'LineWidth', 1.5);
plot(sd_uniform, sa_uniform, 'DisplayName', 'Elastic', 'Color', blue, 'LineWidth', 1.5);
plot(sd_power, sa_power, 'DisplayName', 'Elastic', 'Color', red, 'LineWidth', 1.5);
plot(sd_triangle, scaled_factor * sa, 'DisplayName', 'Elastic', 'Color', green, 'LineWidth', 1.5);
plot(sd_uniform, scaled_factor * sa, 'DisplayName', 'Elastic', 'Color', blue, 'LineWidth', 1.5);
plot(sd_power, scaled_factor * sa, 'DisplayName', 'Elastic', 'Color', red, 'LineWidth', 1.5);

figure;
hold on;
title('ADRS');
xlabel('Roof Displacement(mm)');
ylabel('sa(g)');
plot(sd_triangle * PF1_phi, scaled_factor * sa, 'DisplayName', 'Elastic', 'Color', green, 'LineWidth', 1.5);
plot(sd_uniform * PF1_phi, scaled_factor * sa, 'DisplayName', 'Elastic', 'Color', blue, 'LineWidth', 1.5);
plot(sd_power * PF1_phi, scaled_factor * sa, 'DisplayName', 'Elastic', 'Color', red, 'LineWidth', 1.5);

save('pushover_v2.mat')
