clc; clear; close all;

filenames = [
  "RSN960_NORTHR_LOS000", "RSN960_NORTHR_LOS270", "RSN1602_DUZCE_BOL000", "RSN1602_DUZCE_BOL090", "RSN169_IMPVALL.H_H-DLT262", "RSN169_IMPVALL.H_H-DLT352", "RSN1111_KOBE_NIS000", "RSN1111_KOBE_NIS090", "RSN1158_KOCAELI_DZC180", "RSN1158_KOCAELI_DZC270", "RSN848_LANDERS_CLW-LN", "RSN848_LANDERS_CLW-TR", "RSN752_LOMAP_CAP000", "RSN752_LOMAP_CAP090", "RSN1633_MANJIL_ABBAR--L", "RSN1633_MANJIL_ABBAR--T", "RSN725_SUPER.B_B-POE270", "RSN725_SUPER.B_B-POE360", "RSN1485_CHICHI_TCU045-E", "RSN1485_CHICHI_TCU045-N", "RSN125_FRIULI.A_A-TMZ000", "RSN125_FRIULI.A_A-TMZ270"
];

for index = 1 : length(filenames)
  [ag, time_interval, NPTS, errCode] = parseAT2('../PEERNGARecords_Unscaled/' + filenames(index) + '.AT2');

  if errCode == -1
    error(errCode)
  end

  fprintf('No. %d, Records: %s, PGA: %.3f\n', index, filenames(index), max(abs(ag)));
end

% filename = 'elcentro_EW';



% period = 0 : time_interval : (NPTS - 1) * time_interval;
% ag = Acc;

% tn = 0.01 : 0.01 : 2;
% tn_length = length(tn);
% acceleration = zeros(1, tn_length);

% % time_interval = record_dt;

% for index = 1 : tn_length

%     [~, ~, a_array] = newmark_beta(ag, time_interval, 0.05, tn(index), 'average');

%     acceleration(1, index) = max(abs(a_array));

% end

% tol = eps(0.5);
% fprintf('Records: %s, PGA: %.3f, PGA: %.3f\n', max(abs(ag)), acceleration(1));

% figure;
% plot(tn, acceleration);
% title(filename);
% xlabel('T(sec)');
% ylabel('Sa(g)');

% figure;
% plot(period, ag);
% title(filename);
% xlabel('T(sec)');
% ylabel('Sa(g)');

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
