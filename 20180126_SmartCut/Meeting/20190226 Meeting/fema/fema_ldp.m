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

write_spectrum('test', SDS, SD1, 1)
write_spectrum('design', SDS, SD1, c1(T, T0D) * c2(T, T0D, framing_type, performance_level) * c3(T, theta))
write_spectrum('maximum', SMS, SM1, c1(T, T0M) * c2(T, T0M, framing_type, performance_level) * c3(T, theta))


function [] = write_spectrum(title, Ss, S1, modification_factor)

        [~, sa, tn] = spectrum(Ss, S1, 0.05);

        fileID = fopen([title, ' spectrum.txt'], 'w');
        fprintf(fileID,'%.3f\t%.6f\r\n', [tn; sa * modification_factor]);
        fclose(fileID);

end
