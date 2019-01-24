import os
import sys
import pickle

import pandas as pd

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))


def dataset(filename, storys):
    pkl_file = f'{SCRIPT_DIR}/{filename} for IDA.pkl'

    if not os.path.exists(pkl_file):
        print("Reading excel...")

        read_file = f'{SCRIPT_DIR}/{filename}.xlsx'

        df = pd.read_excel(
            read_file, sheet_name='Story Drifts', header=1, usecols=3, skiprows=[2])

        # convert story label to number
        df = _story2level(df, storys)

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


def _story2level(df, storys):
    for story in storys:
        df.loc[df['Story'] == story, 'StoryLevel'] = storys[story]

    return df
