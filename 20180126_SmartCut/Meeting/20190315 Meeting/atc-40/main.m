clc; clear; close all;

config = Config;

scaled_factor = 0.5 : 9 : 20;

scaled_factor_length = length(scaled_factor);

sd = zeros(1, scaled_factor_length);
sa = zeros(1, scaled_factor_length);

for index = 1 : scaled_factor_length
    [sd(index), sa(index)] = procedure_b(config, 'Triangle', scaled_factor(index));
end

figure;
plot(sd, sa)
