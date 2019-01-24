import matplotlib.pyplot as plt


def peak_interstorey_drift_ratio_versus_storey_level(df, earthquake, earthquakes, scaled_factors):
    sa = earthquakes[earthquake]['sa']

    plt.figure()
    plt.title('Peak interstorey drift ratio versus storey level')
    plt.xlabel(r'Peak interstorey drift ratio $\theta_i$')
    plt.ylabel('Story level')
    plt.xlim(0, 0.03)

    for i in scaled_factors:
        load_case = f'{earthquake}-{i}'
        level_drift = df.loc[df['Load Case/Combo'] == load_case]
        plt.plot(level_drift['Drift'], level_drift['StoryLevel'])

    plt.legend(['%.3fg' % (i * sa) for i in scaled_factors], loc=0)
