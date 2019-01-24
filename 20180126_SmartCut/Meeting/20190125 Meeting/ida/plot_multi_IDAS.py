import matplotlib.pyplot as plt

from single_IDA_points import single_IDA_points


def plot_multi_IDAS(earthquakes, story_drifts, ylim_max=4, xlim_max=0.025, accel_unit='sa'):
    plt.figure()
    plt.title('IDA curves')

    plt.xlabel(r'Maximum interstorey drift ratio, $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(0, xlim_max)
    plt.ylim(0, ylim_max)

    for earthquake in earthquakes:
        drifts, accelerations = single_IDA_points(
            earthquake, earthquakes, story_drifts, accel_unit)

        # xnew = np.linspace(drifts.min(), drifts.max(), 10000000)

        # f = interp1d(drifts, accelerations, kind='cubic')

        if not drifts.size == 0:
            # plt.plot(xnew, f(xnew), label=earthquake, marker='.')
            plt.plot(drifts, accelerations, label=earthquake, marker='.')
        else:
            print(f'{earthquake} is not in data')

    plt.legend(loc='upper left')
