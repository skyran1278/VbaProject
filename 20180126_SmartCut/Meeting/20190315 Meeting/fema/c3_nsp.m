function output = c3_nsp(alpha, R, Te, c3)
    if alpha >= 0
        output = 1.0;
    else
        output = 1.0 + abs(alpha) * (R - 1) ^ (3 / 2) / Te;
    end
    if output > c3
        output = c3;
    end
end
