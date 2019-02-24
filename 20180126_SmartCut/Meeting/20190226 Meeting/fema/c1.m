function output = c1(T, T0)
    if T < 0.1
        output = 1.5;
    elseif T >= T0
        output = 1.0;
    else
        output = (1.0 - 1.5) / (T0 - 0.1) * (T - 0.1) + 1.5;
    end
end
