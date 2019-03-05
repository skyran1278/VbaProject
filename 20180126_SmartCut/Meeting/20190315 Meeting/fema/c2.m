function output = c2(T, T0, framing_type, performance_level)
    % framing_type: 1, 2
    % performace_level: IO, LS, CP
    %
    %    T = 0.1, T >= T0
    %    1   2  , 1    2
    % IO 1.0 1.0, 1.0  1.0
    % LS 1.3 1.0, 1.1  1.0
    % CP 1.5 1.0, 1.2  1.0
    %
    % framing_type 1: Structures in which more than 30% of the story shear at any level is resisted by components or elements whose strength and stiffness may deteriorate during the design earthquake. Such elements and components include: ordinary moment-resisting frames, concentrically-braced frames, frames with partially-restrained connections, tension-only braced frames, unreinforced masonry walls, shear-critical walls and piers, or any combination of the above.
    % framing_type 2: All frames not assigned to Framing Type 1.

        if framing_type == 1
            if performance_level == 'IO'
                output = 1.0;
            elseif performance_level == 'LS'
                if T < 0.1
                    output = 1.3;
                elseif T >= T0
                    output = 1.1;
                else
                    output = (1.1 - 1.3) / (T0 - 0.1) * (T - 0.1) + 1.3;
                end
            elseif performance_level == 'CP'
                if T < 0.1
                    output = 1.5;
                elseif T >= T0
                    output = 1.2;
                else
                    output = (1.2 - 1.5) / (T0 - 0.1) * (T - 0.1) + 1.5;
                end
            end
        elseif framing_type == 2
            output = 1.0;
        end
    end
