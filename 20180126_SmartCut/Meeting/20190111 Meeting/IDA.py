import os
import sys
import pickle

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

storys = {
    'RF': 4,
    '3F': 3,
    '2F': 2,
}

earthquakes = {
    'El Centro': {
        'pga': 0.214,
        'sa': 0.414
    },
    'TAP010': {
        'pga': 0.117,
        'sa': 0.171,
    },
    'TCU052': {
        'pga': 0.447,
        'sa': 0.683
    },
    'TCU067': {
        'pga': 0.498,
        'sa': 1.234
    },
    'TCU068': {
        'pga': 0.511,
        'sa': 1.383
    },
}


def dataset():
    filename = '20190107 story drifts.xlsx'

    read_file = f'{SCRIPT_DIR}/{filename}'

    df = pd.read_excel(
        read_file, sheet_name='Story Drifts', header=1, usecols=3, skiprows=[2])

    # convert story label to number
    df = story2level(df, storys)

    # delete max and min string
    df.loc[:, 'Load Case/Combo'] = df['Load Case/Combo'].str[:-4]

    df = df.assign(
        StoryAndCase=lambda x: x['Story'] + ' ' + x['Load Case/Combo'])

    # combine max min
    df = df.groupby('StoryAndCase', as_index=False, sort=False).agg('max')

    df['Load Case'], df['Scaled Factors'] = df['Load Case/Combo'].str.rsplit(
        '-', 1).str

    return df


def story2level(df, storys):
    for story in storys:
        df.loc[df['Story'] == story, 'StoryLevel'] = storys[story]

    return df


story_drifts = dataset()
print(story_drifts.head())


def single_IDA_points(earthquake, earthquakes, story_drifts, accel_unit='sa'):
    accel = earthquakes[earthquake][accel_unit]

    earthquake_drift = story_drifts.loc[story_drifts['Load Case']
                                        == earthquake, :]

    max_drift = earthquake_drift.groupby(
        'Scaled Factors', as_index=False, sort=False)['Drift'].max()

    max_drift.loc[:, 'Scaled Factors'] = max_drift.loc[:,
                                                       'Scaled Factors'].astype('float64') * accel

    return max_drift['Drift'], max_drift['Scaled Factors']


def plot_single_IDA(earthquake, earthquakes, story_drifts, xlim_max=0.025, accel_unit='sa'):
    plt.figure()
    plt.title('Single IDA curve')

    plt.xlabel(r'Maximum interstorey drift ratio $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim((0, xlim_max))

    drifts, accelerations = single_IDA_points(
        earthquake, earthquakes, story_drifts, accel_unit)

    if not drifts.empty:
        plt.plot(drifts, accelerations)
    else:
        print(f'{earthquake} is not in data')


def plot_multi_IDAS(earthquakes, story_drifts, xlim_max=0.025, accel_unit='sa'):
    plt.figure()
    plt.title('IDA curves')

    plt.xlabel(r'Maximum interstorey drift ratio $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim((0, xlim_max))

    for earthquake in earthquakes:
        drifts, accelerations = single_IDA_points(
            earthquake, earthquakes, story_drifts, accel_unit)

        if not drifts.empty:
            plt.plot(drifts, accelerations, label=earthquake)
        else:
            print(f'{earthquake} is not in data')

    plt.legend(loc='upper right')


plot_single_IDA('TCU067', earthquakes, story_drifts)
plot_multi_IDAS(earthquakes, story_drifts)
plt.show()


def peak_interstorey_drift_ratio_versus_storey_level(df, earthquake, earthquakes, scaled_factors):
    sa = earthquakes[earthquake]['sa']

    plt.figure()
    plt.title('Peak interstorey drift ratio versus storey level')
    plt.xlabel(r'Peak interstorey drift ratio $\theta_i$')
    plt.ylabel('Story level')
    plt.xlim((0, 0.03))

    for i in scaled_factors:
        load_case = f'{earthquake}-{i}'
        level_drift = df.loc[df['Load Case/Combo'] == load_case]
        plt.plot(level_drift['Drift'], level_drift['StoryLevel'])

    plt.legend(['%.3fg' % (i * sa) for i in scaled_factors], loc=0)
