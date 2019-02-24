function output = c3(T, theta)
    if theta <= 0.1
        output = 1.0;
    else
        output = 1 + 5 * (theta - 1) / T;
    end
end
