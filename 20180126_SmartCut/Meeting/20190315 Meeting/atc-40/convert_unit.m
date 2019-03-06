function factor = convert_unit(unit)
    if unit == "gal2m"
        factor = 0.01;
    elseif unit == "gal2g"
        factor = 0.00101972; % 1 / 980.665
    elseif unit == "gal2mm"
        factor = 10;
    end
end
