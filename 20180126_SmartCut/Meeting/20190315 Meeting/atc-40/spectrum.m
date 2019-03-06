function [sd, sa, tn] = spectrum(time_history_filename, damping_ratio, scaled_factor)

    narginchk(1, 3);

    if nargin <= 2
        scaled_factor = 1.0;
    end

    if nargin == 1
        damping_ratio = 0.05;
    end

    period = filename_to_array(time_history_filename, 2, 1); % s
    ag = filename_to_array(time_history_filename, 2, 2) * scaled_factor; % g

    tn = 0.01 : 0.01 : 5;

    tn_length = length(tn);
    sd = zeros(1, tn_length);
    sa = zeros(1, tn_length);

    time_interval = period(2) - period(1);

    for index = 1 : tn_length

        [u_array, ~, a_array] = newmark_beta(ag, time_interval, damping_ratio, tn(index), 'average');

        sd(1, index) = max(abs(u_array));
        sa(1, index) = max(abs(a_array));

    end

    sd = sd * 981; % g to mm

end
