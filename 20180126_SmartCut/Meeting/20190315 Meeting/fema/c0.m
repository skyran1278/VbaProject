function output = c0(stories_number)
    if stories_number == 1
        output = 1.0;
    elseif stories_number == 2
        output = 1.2;
    elseif stories_number == 3
        output = 1.3;
    elseif stories_number == 4
        output = 1.35;
    elseif stories_number == 5
        output = 1.4;
    elseif stories_number >= 10
        output = 1.5;
    else
        output = (1.5 - 1.4) / (10 - 5) * (stories_number - 5) + 1.4;
    end
end
