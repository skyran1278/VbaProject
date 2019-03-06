clc; clear; close all;

config = Config;

scaled_factor = 0.5 : 0.5 : 15;

scaled_factor_length = length(scaled_factor);

sd = NaN(1, scaled_factor_length);
sa = NaN(1, scaled_factor_length);

for index = 1 : scaled_factor_length
    [sd(index), sa(index)] = procedure_b(config, 'Triangle', scaled_factor(index));
end

figure;
plot(sd, sa)

for index = 1 : scaled_factor_length
    [sd(index), sa(index)] = procedure_b(config, 'Uniform', scaled_factor(index));
end

figure;
plot(sd, sa)

for index = 1 : scaled_factor_length
    [sd(index), sa(index)] = procedure_b(config, 'Power', scaled_factor(index));
end

figure;
plot(sd, sa)