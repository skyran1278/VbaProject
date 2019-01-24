import matplotlib.pyplot as plt

from single_IDA_points import single_IDA_points


def plot_single_IDA(earthquake, earthquakes, story_drifts, ylim_max=4, xlim_max=0.025, accel_unit='sa'):
    plt.figure()
    plt.title('Single IDA curve')

    plt.xlabel(r'Maximum interstorey drift ratio $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim(0, xlim_max)
    plt.ylim(0, ylim_max)

    drifts, accelerations = single_IDA_points(
        earthquake, earthquakes, story_drifts, accel_unit)

    if not drifts.size == 0:
        plt.plot(drifts, accelerations)
    else:
        print(f'{earthquake} is not in data')
