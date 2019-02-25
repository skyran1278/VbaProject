function output = c1_nsp(Te, T0, R, c1)
    if Te >= T0
        output = 1.0;
    else
        output = (1.0 + (R - 1) * T0 / Te) / R;
    end

    if output > c1
        output = c1;
    end
end
