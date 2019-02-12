import matplotlib.pyplot as plt

from interp_IDAS import interp_IDAS


def plot_fractiles_log(earthquakes, story_drifts, ylim_max=4, xlim_max=0.025, accel_unit='sa'):
    interpolation_x, interpolation_y = interp_IDAS(
        earthquakes, story_drifts, accel_unit)

    plt.figure()

    plt.title('16%, 50% and 84% fractiles')

    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(10**-4, xlim_max)

    plt.loglog(interpolation_x.quantile(0.5, axis=1, interpolation='nearest'),
               interpolation_y, label='50%')
    plt.loglog(interpolation_x.quantile(0.16, axis=1, interpolation='nearest'),
               interpolation_y, label='16%')
    plt.loglog(interpolation_x.quantile(0.84, axis=1, interpolation='nearest'),
               interpolation_y, label='84%')

    plt.legend(loc='upper left')
    plt.grid(True, which="both")


def plot_fractiles(earthquakes, story_drifts, ylim_max=4, xlim_max=0.025, accel_unit='sa'):
    interpolation_x, interpolation_y = interp_IDAS(
        earthquakes, story_drifts, accel_unit)

    plt.figure()

    plt.title('16%, 50% and 84% fractiles')

    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(0, xlim_max)
    plt.ylim(0, ylim_max)

    # plot 50% 16% 84%
    plt.plot(interpolation_x.quantile(0.5, axis=1, interpolation='nearest'),
             interpolation_y, label='50%')
    plt.plot(interpolation_x.quantile(0.16, axis=1, interpolation='nearest'),
             interpolation_y, label='16%')
    plt.plot(interpolation_x.quantile(0.84, axis=1, interpolation='nearest'),
             interpolation_y, label='84%')

    plt.legend(loc='upper left')
