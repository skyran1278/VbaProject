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

scaled_factors = [1, 2, 4, 5, 8, 10, 15, 16, 20]

earthquakes = {
    'EL Centro': {
        'pga': 0.214,
        'sa': 0.415
    },
    'chichi_TAP010': {
        'pga': 0.117,
        'sa': 0.172
    },
}


def dataset():
    filename = '20181229 story drift.xlsx'

    read_file = f'{SCRIPT_DIR}/{filename}'

    df = pd.read_excel(
        read_file, sheet_name='Story Drifts', header=1, usecols=3, skiprows=[2])

    df = df.assign(StoryLevel=None)

    df = story2level(df, storys)

    df = df.assign(
        StoryAndCase=lambda x: x['Story'] + ' ' + x['Load Case/Combo'])

    df.loc[:, 'StoryAndCase'] = df['StoryAndCase'].str[:-4]

    df = df.groupby('StoryAndCase', as_index=False, sort=False).agg('max')

    df.loc[:, 'Load Case/Combo'] = df['Load Case/Combo'].str[:-4]

    return df


def story2level(df, storys):
    for story in storys:
        df.loc[df['Story'] == story, 'StoryLevel'] = storys[story]

    return df


story_drifts = dataset()
print(story_drifts.head())


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


def single_IDA_curve_versus_static_pushover(df, earthquake, earthquakes, scaled_factors):
    sa = earthquakes[earthquake]['sa']

    drifts = []
    sas = []

    plt.figure()
    plt.title('Single IDA curve versus Static Pushover')
    plt.xlabel(r'Maximum interstorey drift ratio $\theta_i$')
    plt.ylabel(r'"first-mode"spectral acceleration $S_a(T_1$, 5%)')
    # plt.xlim((0, 0.3))

    max_drift = df.groupby(
        'Load Case/Combo', as_index=False, sort=False).agg('max')

    for i in scaled_factors:
        load_case = f'{earthquake}-{i}'
        drift = max_drift.loc[df['Load Case/Combo']
                              == load_case, 'Drift'].iat[0]
        sas.append(sa * i)
        drifts.append(drift)

    plt.plot(drifts, sas)

    # plt.legend(['%.3fg' % (i * sa) for i in scaled_factors], loc=0)


single_IDA_curve_versus_static_pushover(
    story_drifts, 'chichi_TAP010', earthquakes, scaled_factors)
peak_interstorey_drift_ratio_versus_storey_level(
    story_drifts, 'chichi_TAP010', earthquakes, [1, 2, 4, 5])
plt.show()
