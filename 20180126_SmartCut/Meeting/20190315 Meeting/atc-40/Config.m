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
        ];

        power = [
            0	0
            8.904	0.280781
            19.535	0.452255
            31.077	0.48542
            35.436	0.498543
            38.481	0.501705
            50.036	0.50453
            61.594	0.508297
            73.152	0.512204
            84.712	0.516698
            96.272	0.52131
            104.945	0.525048
        ];

        uniform = [
            0	0
            9.396	0.38995
            15.281	0.519193
            26.946	0.567329
            32.438	0.596899
            35.346	0.605905
            46.583	0.627426
            58.142	0.63554
            69.704	0.643323
            81.266	0.650824
            92.83	0.658184
            97.418	0.661245
        ];

    end

    methods
        function [sd, sa] = load_pattern(obj, name)
            if name == "triangle"
                sd = obj.triangle(:, 1).';
                sa = obj.triangle(:, 2).';
            elseif name == "uniform"
                sd = obj.uniform(:, 1).';
                sa = obj.uniform(:, 2).';
            elseif name == "power"
                sd = obj.power(:, 1).';
                sa = obj.power(:, 2).';
            end
        end
    end

end
