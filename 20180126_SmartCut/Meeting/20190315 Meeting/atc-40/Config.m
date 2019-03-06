classdef Config

    properties
        filename = 'TAP010 unit_g';

        structural_behavior_type = 'A';

        triangle = [
            0	0
            9.104	0.318839
            17.83	0.478117
            29.423	0.516952
            34.175	0.536776
            41.672	0.547048
            53.236	0.552322
            64.796	0.557274
            76.36	0.5628
            87.922	0.568285
            99.485	0.573934
            101.957	0.575255
            102.173	0.335883
            109.015	0.384201
            112.643	0.396123
            116.507	0.397642
        ];
    end

    methods
        function [sd, sa] = load_pattern(obj, name)
            if name == "Triangle"
                sd = obj.triangle(:, 1).';
                sa = obj.triangle(:, 2).';
            end
        end
    end

end
