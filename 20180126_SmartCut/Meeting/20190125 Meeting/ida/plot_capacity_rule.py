import numpy as np
import matplotlib.pyplot as plt

from scipy.interpolate import spline

from interp_IDAS import interp_IDAS


def plot_IM_rule(story_drifts, scaled_earthquake, ylim_max=4, xlim_max=0.025, accel_unit='sa', C_IM=0.2):
    # interpolation_x, interpolation_y = interp_IDAS(
    #     earthquakes, story_drifts, accel_unit)

    # plt.figure()

    # for earthquake in interpolation_x:
    #     plt.plot(interpolation_x[earthquake],
    #              interpolation_y, label=earthquake)

    # interpolation_x = interpolation_x.quantile(
    #     0.5, axis=1, interpolation='nearest').values

    stiffness = np.gradient(scaled_earthquake, story_drifts)

    plt.figure()

    plt.plot(story_drifts, stiffness)

    # print(stiffness)

    # initial_stiffness = (
    #     scaled_earthquake[1] - scaled_earthquake[0]) / (interpolation_x[1] - interpolation_x[0])

    # plt.legend(loc='upper right')

    plt.figure()

    plt.title('IM-base rule')

    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(0, xlim_max)
    plt.ylim(0, ylim_max)

    # plot 50%
    plt.plot(story_drifts, scaled_earthquake)
    plt.plot([0, xlim_max], [C_IM, C_IM], '--')


def plot_DM_rule(story_drifts, scaled_earthquake, ylim_max=4, xlim_max=0.025, accel_unit='sa', C_DM=0.02):
    # interpolation_x, interpolation_y = interp_IDAS(
    #     earthquakes, story_drifts, accel_unit)

    # plt.figure()

    # for earthquake in interpolation_x:
    #     plt.plot(interpolation_x[earthquake],
    #              interpolation_y, label=earthquake)

    # interpolation_x = interpolation_x.quantile(
    #     0.5, axis=1, interpolation='nearest').values

    # stiffness = np.gradient(interpolation_y, interpolation_x)

    # initial_stiffness = (
    #     interpolation_y[1] - interpolation_y[0]) / (interpolation_x[1] - interpolation_x[0])

    # plt.legend(loc='upper right')
    plt.figure()
    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')
    plt.title('DM-base rule')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(0, xlim_max)
    plt.ylim(0, ylim_max)

    # plot 50%
    plt.plot(story_drifts, scaled_earthquake)
    plt.plot([C_DM, C_DM], [0, ylim_max], '--')
