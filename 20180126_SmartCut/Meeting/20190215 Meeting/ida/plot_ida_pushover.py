import matplotlib.pyplot as plt

from plot_multi_IDAS import plot_multi_IDAS
from interp_IDAS import interp_IDAS


def plot_multi_idas_and_pushover(earthquakes, multi_story_drifts, pushover, loadcases, ylim_max=None, xlim_max=0.025):
    plot_multi_IDAS(earthquakes, multi_story_drifts,
                    ylim_max, xlim_max, accel_unit='sa')
    pushover.plot(loadcases, marker='.')
    plt.title('IDA versus Static Pushover for a 3-storey moment resisting frame')
    plt.legend(loc='upper left')


# def plot_median_idas_and_pushover(earthquakes, multi_story_drifts, pushover, ylim_max=None, xlim_max=0.025):
#     plt.figure()
#     plt.title('IDA versus Static Pushover for a 3-storey moment resisting frame')

#     plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')
#     plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')

#     if xlim_max is not None:
#         plt.xlim(0, xlim_max)
#     if ylim_max is not None:
#         plt.ylim(0, ylim_max)

#     multi_x, multi_y = _median_ida_points(
#         earthquakes, multi_story_drifts, accel_unit='sa')
#     plt.plot(multi_x, multi_y, label='median IDA curve', marker='.')

#     pushover.plot_all_loadcases()

#     plt.legend(loc='upper left')


def plot_median_idas_and_pushover(earthquakes, multi_story_drifts, pushover, loadcases, ylim_max=None, xlim_max=0.025):
    plt.figure()
    plt.title('IDA versus Static Pushover for a 3-storey moment resisting frame')

    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')
    plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')

    if xlim_max is not None:
        plt.xlim(0, xlim_max)
    if ylim_max is not None:
        plt.ylim(0, ylim_max)

    multi_x, multi_y = _median_ida_points(
        earthquakes, multi_story_drifts, accel_unit='sa')
    plt.plot(multi_x, multi_y, label='median IDA curve')

    pushover.plot(loadcases, marker='.')

    plt.legend(loc='upper left')


def _median_ida_points(earthquakes, multi_story_drifts, accel_unit='sa'):
    multi_x, multi_y = interp_IDAS(
        earthquakes, multi_story_drifts, accel_unit=accel_unit)
    multi_x = multi_x.quantile(
        0.5, axis=1, interpolation='nearest').values

    return multi_x, multi_y
