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
        'scaled_factors': [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2, 2.1,
                           2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9, 3, 4, 5, 5.1,  6, 7, 8, 9, 10]
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


def plot_single_IDA(earthquake, earthquakes, df, accel_unit='sa', xlim_max=0.15):
    plt.figure()
    plt.title('Single IDA curve')

    plt.xlabel(r'Maximum interstorey drift ratio $\theta_{max}$')

    if accel_unit == 'sa':
        plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)(g)')
    elif accel_unit == 'pga':
        plt.ylabel('Peak ground acceleration PGA(g)')

    plt.xlim((0, xlim_max))

    drifts, accelerations = single_IDA_points(
        earthquake, earthquakes, df, accel_unit)

    plt.plot(drifts, accelerations)


plot_single_IDA('TAP010', earthquakes, story_drifts)
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
