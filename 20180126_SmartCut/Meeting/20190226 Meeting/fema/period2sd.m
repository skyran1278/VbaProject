function sd = period2sd(period, acceleration)
    sd = (period .^ 2) / (4 * pi .^ 2) .* acceleration * 981;
end
