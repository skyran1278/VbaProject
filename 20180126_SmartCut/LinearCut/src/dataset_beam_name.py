""" load user defined beam name
"""
import os
import pickle

import pandas as pd


def _load_name(read_file):
    return pd.read_excel(read_file, sheet_name='梁名編號', index_col=[0, 1], usecols=[1, 2, 3, 4])


def _init_pkl(read_file, save_file):
    dataset = _load_name(read_file)

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_beam_name(read_file, save_file):
    """ load beam name
    """
    if not os.path.exists(save_file):
        _init_pkl(read_file, save_file)

    with open(save_file, 'rb') as f:
        return pickle.load(f)


if __name__ == "__main__":
    from const import const
    BEAM_NAME_PATH = const['beam_name_path']

    READ_FILE = f'{BEAM_NAME_PATH}'
    SAVE_FILE = f'{BEAM_NAME_PATH}.pkl'

    _init_pkl(READ_FILE, SAVE_FILE)
    DATASET = load_beam_name(READ_FILE, SAVE_FILE)
    print(DATASET.head())
    # print(dataset['樓層'])
    # print(list(dataset))