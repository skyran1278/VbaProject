function [sd, sa, tn] = spectrum(SDS, SD1, damping_ratio)

    [BS, B1] = damping_factor(damping_ratio);

    T0 = SD1 * BS / (SDS * B1);
    % T0 = SD1 / (SDS);

    tn = 0 : 0.1 : 5;

    tn_length = length(tn);
    sa = zeros(1, tn_length);

    for index = 1 : tn_length
        T = tn(index);

        if T <= 0.2 * T0
            sa(1, index) = SDS * (0.4 + (1 / BS - 0.4) * T / (0.2 * T0));

        elseif T > (0.2 * T0) && T <= T0
            sa(1, index) = SDS / BS;

        elseif T > T0 && T <= 2.5 * T0
            sa(1, index) = SD1 / (B1 * T);

        else
            sa(1, index) = 0.4 * SDS / BS;

        end
    end

    sd = period2sd(tn, sa);

end
