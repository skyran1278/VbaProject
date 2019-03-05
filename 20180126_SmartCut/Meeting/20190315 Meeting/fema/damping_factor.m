function [BS, B1] = damping_factor(damping_ratio)
    if damping_ratio <= 0.02
        BS = 0.8;
        B1 = 0.8;

    elseif 0.02 < damping_ratio && damping_ratio <= 0.05
        BS = (damping_ratio - 0.02) / 0.03 * (1 - 0.8) + 0.8;
        B1 = (damping_ratio - 0.02) / 0.03 * (1 - 0.8) + 0.8;

    elseif 0.05 < damping_ratio && damping_ratio <= 0.1
        BS = (damping_ratio - 0.05) / 0.05 * (1.33 - 1) + 1;
        B1 = (damping_ratio - 0.05) / 0.05 * (1.25 - 1) + 1;

    elseif 0.1 < damping_ratio && damping_ratio <= 0.2
        BS = (damping_ratio - 0.1) / 0.1 * (1.6 - 1.33) + 1.33;
        B1 = (damping_ratio - 0.1) / 0.1 * (1.5 - 1.25) + 1.25;

    elseif 0.2 < damping_ratio && damping_ratio <= 0.3
        BS = (damping_ratio - 0.2) / 0.1 * (1.79 - 1.6) + 1.6;
        B1 = (damping_ratio - 0.2) / 0.1 * (1.63 - 1.5) + 1.5;

    elseif 0.3 < damping_ratio && damping_ratio <= 0.4
        BS = (damping_ratio - 0.3) / 0.1 * (1.87 - 1.79) + 1.79;
        B1 = (damping_ratio - 0.3) / 0.1 * (1.7 - 1.63) + 1.63;

    elseif 0.4 < damping_ratio && damping_ratio <= 0.5
        BS = (damping_ratio - 0.4) / 0.1 * (1.93 - 1.87) + 1.87;
        B1 = (damping_ratio - 0.4) / 0.1 * (1.75 - 1.7) + 1.7;

    elseif 0.5 < damping_ratio
        BS = 1.93;
        B1 = 1.75;

    end
end
