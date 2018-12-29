clc; clear; close all;

filename = 'elcentro_EW';

period = filename_to_array(filename, 2, 1);
ag = filename_to_array(filename, 2, 2);

% ag = ag / max(abs(ag)) * 0.4;

tn = 0.001 : 0.001 : 1;
tn_length = length(tn);
acceleration = zeros(1, tn_length);

time_interval = period(2) - period(1);

for index = 1 : tn_length

    [~, ~, a_array] = newmark_beta(ag, time_interval, 0.05, tn(index), 'average');

    acceleration(1, index) = max(abs(a_array));

end

fprintf('PGA: %.3f, PGA: %.3f, Sa: %.3f\n', max(abs(ag)), acceleration(1), acceleration(345));

figure;
plot(tn, acceleration);
title(filename);
xlabel('T(sec)');
ylabel('Sa(g)');

% figure;
% hold on;
% title(filename);
% xlabel('T(sec)');
% ylabel('Sa(g)');

% for intensity = 0.5 : 0.5 : 3

%     for index = 1 : tn_length

%         scaled_ag = intensity * ag;

%         [~, ~, a_array] = newmark_beta(scaled_ag, time_interval, 0.05, tn(index), 'average');

%         acceleration(1, index) = max(abs(a_array));

%     end

%     fprintf('Intensity: %.1f, PGA: %.3f, PGA: %.3f, Sa: %.3f, PGA Scaled: %.3f, Sa Scaled: %.3f\n', intensity, max(scaled_ag), acceleration(1), acceleration(tn == 0.295), max(scaled_ag) / intensity, acceleration(tn == 0.295) / intensity);

%     plot(tn, acceleration);

% end



% fileID = fopen('ELC.txt', 'w');
% fprintf(fileID, '%f \r\n', ag);
% fclose(fileID);
% clc; clear; close all;

% ag = filename_to_array('I-ELC270_gal_l00Hz', 2, 2);

% tn = 0.01 : 0.01 : 5;
% tn_length = length(tn);
% acceleration = zeros(1, tn_length);

% for index = 1 : tn_length

%     [~, ~, a_array] = newmark_beta(ag, 0.01, 0.05, tn(index), 'average');

%     acceleration(1, index) = max(abs(a_array));

% end

% acceleration_normal = acceleration / acceleration(1, 1) * 0.4;
% acceleration_normal(tn == 2.5)
% figure;
% plot(tn, acceleration_normal);
% xlabel('T(sec)');
% ylabel('SaD(g)');
