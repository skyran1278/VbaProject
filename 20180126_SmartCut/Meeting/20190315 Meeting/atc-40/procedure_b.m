clc; clear; close all;


config = Config;

[capacity_sd, capacity_sa] = config.load_pattern("Triangle");

[elastic_sd, elastic_sa] = get_elastic_line(capacity_sd, capacity_sa);

[demand_sd, demand_sa, ~] = spectrum(config.filename);

[d_star, a_star] = get_star_point(elastic_sd, elastic_sa, capacity_sd, capacity_sa, demand_sd, demand_sa);

[dy, ay] = get_yielding_point(elastic_sd, elastic_sa, capacity_sd, capacity_sa, d_star, a_star);

[beta_eff, dpi] = get_beff_and_dpi(demand_sd, d_star, a_star, dy, ay, config.structural_behavior_type);

[single_demand_sd, single_demand_sa] = get_single_demand(config, beta_eff, dpi);

[sd, sa] = get_performance_point(single_demand_sd, single_demand_sa, elastic_sd, elastic_sa, dy, ay, d_star, a_star, capacity_sd);

figure;
hold on;
plot(capacity_sd, capacity_sa);
plot(elastic_sd, elastic_sa);
plot(demand_sd, demand_sa);
plot(single_demand_sd, single_demand_sa);
plot(d_star, a_star, 'bo');
plot(dy, ay, 'ko');
title('');
xlabel('');
ylabel('');
axis([0 inf 0 0.6]);

function y = linear_interpolate(x, x1, x2, y1, y2)
    y = (y2 - y1) / (x2 - x1) * (x - x1) + y1;
end

function [elastic_sd, elastic_sa] = get_elastic_line(sd, sa)
    elastic_sd = [sd(1), sd(2), sd(end)];
    elastic_sa_end = linear_interpolate(sd(end), sd(1), sd(2), sa(1), sa(2));
    elastic_sa = [sa(1), sa(2), elastic_sa_end];
end

function [d_star, a_star] = get_star_point(elastic_sd, elastic_sa, capacity_sd, capacity_sa, demand_sd, demand_sa)
    point_temp = InterX([elastic_sd; elastic_sa], [demand_sd; demand_sa]);
    point_star = InterX([capacity_sd; capacity_sa], [[point_temp(1, :), point_temp(1, :)]; [0, point_temp(2, :)]]);

    d_star = point_star(1, :);
    a_star = point_star(2, :);

    if size(point_temp, 2) ~= 1 || size(point_star, 2) ~= 1
        print('star point not only 1 intersection')
    end

end

function [dy, ay] = get_yielding_point(elastic_sd, elastic_sa, capacity_sd, capacity_sa, d_star, a_star)
    bilinear_area = 0;
    dy = elastic_sd(1);
    ay = elastic_sa(1);

    capacity_area = trapz([capacity_sd(capacity_sd < d_star), d_star], [capacity_sa(capacity_sd < d_star), a_star]);

    while abs(bilinear_area - capacity_area) > 1e-4
        dy = dy + 1e-4;
        ay = linear_interpolate(dy, elastic_sd(1), elastic_sd(2), elastic_sa(2), elastic_sa(1));

        bilinear_area = trapz([elastic_sd(1), dy, d_star], [elastic_sa(1), ay, a_star]);
    end
end

function [beta_eff, dpi] = get_beff_and_dpi(demand_sd, d_star, a_star, dy, ay, structural_behavior_type)
    % dpi have to intersection with demand curve, so choose max demand curve 5%
    dpi = dy : 0.1 : demand_sd(end); % matrix

    % api = (a_star - ay) * (dpi - dy) / (d_star - dy) + ay; % matrix
    api = linear_interpolate(dpi, dy, d_star, ay, a_star); % matrix

    beta_0 = 63.7 * (ay * dpi - dy * api) ./ (api .* dpi); % matrix

    if structural_behavior_type == 'A'
        kappa = get_kappa(beta_0, dpi, api, dy, ay);
    end

    beta_eff = kappa .* beta_0 + 5;

end

function kappa = get_kappa(beta_0, dpi, api, dy, ay)
    beta_0_length = length(beta_0);

    kappa = zeros(1, beta_0_length);

    for index = 1 : beta_0_length

        if beta_0(index) <= 16.5
            kappa(index) = 1.0;
        else
            kappa(index) = 1.13 - 0.51 * (ay * dpi(index) - dy * api(index)) / (api(index) * dpi(index));
        end

    end

end

function [sd, sa] = get_single_demand(config, beta_eff, dpi)
    beta_eff_length = length(beta_eff);

    sd = NaN(1, beta_eff_length);
    sa = NaN(1, beta_eff_length);

    for index = 1 : beta_eff_length

        [demand_sd, demand_sa, ~] = spectrum(config.filename, beta_eff(index) / 100);

        single_demand = InterX([demand_sd; demand_sa], [[dpi(index), dpi(index)]; [0, max(demand_sa)]]);

        if ~isempty(single_demand)
            % may be multi intersection, so select max sa
            sd(index) = max(single_demand(1, :));
            sa(index) = max(single_demand(2, :));
        end

    end

end

function [sd, sa] = get_performance_point(single_demand_sd, single_demand_sa, elastic_sd, elastic_sa, dy, ay, d_star, a_star, capacity_sd)
    bilinear_sd = [elastic_sd(1), dy, d_star, capacity_sd];
    bilinear_sa_end = linear_interpolate(capacity_sd, dy, d_star, ay, a_star);
    bilinear_sa = [elastic_sa(1), ay, a_star, bilinear_sa_end];

    point_temp = InterX([single_demand_sd; single_demand_sa], [bilinear_sd; bilinear_sa]);

    if size(point_temp, 2) ~= 1
        print('performance point not only 1 intersection')
    end

    sd = point_temp(1, :);
    sa = point_temp(2, :);

end
