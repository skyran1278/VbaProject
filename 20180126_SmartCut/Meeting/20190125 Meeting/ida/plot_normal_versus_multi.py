import matplotlib.pyplot as plt

from interp_IDAS import interp_IDAS


def plot_normal_versus_multi_log(earthquakes, normal_story_drifts, multi_story_drifts, ylim_max=4, xlim_max=0.025, accel_unit='sa'):
    normal_x, normal_y = interp_IDAS(
        earthquakes, normal_story_drifts, accel_unit)

    multi_x, multi_y = interp_IDAS(
        earthquakes, multi_story_drifts, accel_unit)

    plt.figure()

    plt.title('Normal versus Multi IDA for a 3-storey MRF')

    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(10**-4, xlim_max)

    plt.loglog(normal_x.quantile(0.5, axis=1, interpolation='nearest'),
               normal_y, label='50%')
    plt.loglog(multi_x.quantile(0.5, axis=1, interpolation='nearest'),
               multi_y, label='50%')

    plt.legend(loc='upper left')
    plt.grid(True, which="both")


def plot_normal_versus_multi(earthquakes, normal_story_drifts, multi_story_drifts, ylim_max=4, xlim_max=0.025, accel_unit='sa'):
    normal_x, normal_y = interp_IDAS(
        earthquakes, normal_story_drifts, accel_unit)

    multi_x, multi_y = interp_IDAS(
        earthquakes, multi_story_drifts, accel_unit)

    plt.figure()

    plt.title('Normal versus Multi IDA for a 3-storey MRF')

    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(0, xlim_max)
    plt.ylim(0, ylim_max)

    plt.plot(normal_x.quantile(0.5, axis=1, interpolation='nearest'),
             normal_y, label='50%')
    plt.plot(multi_x.quantile(0.5, axis=1, interpolation='nearest'),
             multi_y, label='50%')

    plt.legend(loc='upper left')
