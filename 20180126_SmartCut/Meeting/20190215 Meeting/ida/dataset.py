""" reading excel
"""
import os
import pickle

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def dataset_ida_storydrifts(filename, stories):
    """ get time history max story drift
    """
    pkl_file = f'{SCRIPT_DIR}/{filename} for IDA.pkl'

    if not os.path.exists(pkl_file):
        print("Reading excel...")

        read_file = f'{SCRIPT_DIR}/{filename}.xlsx'

        df = pd.read_excel(
            read_file, sheet_name='Story Drifts', header=1, usecols=3, skiprows=[2])

        # convert story label to number
        df = _story2level(df, stories)

        # delete max and min string
        df.loc[:, 'Load Case/Combo'] = df['Load Case/Combo'].str[:-4]

        df = df.assign(
            StoryAndCase=lambda x: x['Story'] + ' ' + x['Load Case/Combo'])

        # combine max min
        df = df.groupby('StoryAndCase', as_index=False, sort=False).agg('max')

        df['Load Case'], df['Scaled Factors'] = df['Load Case/Combo'].str.rsplit(
            '-', 1).str

        print("Creating pickle file ...")
        with open(pkl_file, 'wb') as f:
            pickle.dump(df, f, True)
        print("Done!")

    with open(pkl_file, 'rb') as f:
        df = pickle.load(f)

    return df


def dataset_pushover_storydrifts(filename, stories):
    """
    get every step pushover story drifts
    """
    pkl_file = f'{SCRIPT_DIR}/{filename} for pushover.pkl'

    if not os.path.exists(pkl_file):
        print("Reading excel...")

        read_file = f'{SCRIPT_DIR}/{filename}.xlsx'

        df = pd.read_excel(
            read_file, sheet_name='Story Drifts', header=1, usecols=3, skiprows=[2])

        # convert story label to number
        df = _story2level(df, stories)

        # StoryAndCase = Story + Load Case/Combo
        df = df.assign(
            StoryAndCase=lambda x: x['Story'] + ' ' + x['Load Case/Combo'])

        # split Load Case/Combo to load case and step
        df['Load Case'], df['Step'] = df['Load Case/Combo'].str.rsplit(
            ' ', 1).str

        print("Creating pickle file ...")
        with open(pkl_file, 'wb') as f:
            pickle.dump(df, f, True)
        print("Done!")

    with open(pkl_file, 'rb') as f:
        df = pickle.load(f)

    return df


def get_max_story_drift(df):
    """
    condense story drift to max drift
    """
    return df.groupby(
        'Load Case/Combo', as_index=False, sort=False)['Drift'].max()


def dataset_pushover_baseshear(filename):
    """
    get pushover base shear and acceleration
    """
    pkl_file = f'{SCRIPT_DIR}/{filename} for pushover.pkl'

    if not os.path.exists(pkl_file):
        print("Reading excel...")

        read_file = f'{SCRIPT_DIR}/{filename}.xlsx'

        df = pd.read_excel(
            read_file, sheet_name='Base Reactions', header=1, usecols=3, skiprows=[2])

        df['Load Case'], df['Step'] = df['Load Case/Combo'].str.rsplit(
            ' ', 1).str

        df['Accel'] = np.abs(df['FX'] / df['FZ'])

        print(df.head())

        print("Creating pickle file ...")
        with open(pkl_file, 'wb') as f:
            pickle.dump(df, f, True)
        print("Done!")

    with open(pkl_file, 'rb') as f:
        df = pickle.load(f)

    return df


def _story2level(df, stories):
    for story in stories:
        df.loc[df['Story'] == story, 'StoryLevel'] = stories[story]

    return df


def _main():
    stories = {
        'RF': 4,
        '3F': 3,
        '2F': 2,
    }

    pushouver_story_drifts = dataset_pushover_storydrifts(
        '20190212 pushover story drifts', stories)

    get_max_story_drift(pushouver_story_drifts)

    pushouver_base_shear = dataset_pushover_baseshear(
        '20190212 pushover base shear')


if __name__ == "__main__":
    _main()
