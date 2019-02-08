import os
import sys
import pickle

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

read_file = f'{SCRIPT_DIR}/../out/first_run.xlsx'
save_file = f'{SCRIPT_DIR}/../temp/beam_name.pkl'


def _load_name(read_file):
    return pd.read_excel(read_file, sheet_name='梁名編號', index_col=[0, 1], usecols=[1, 2, 3, 4])


def _init_pkl(read_file, save_file):
    dataset = _load_name(read_file)

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_beam_name(read_file=read_file, save_file=save_file):
    if not os.path.exists(save_file):
        _init_pkl(read_file, save_file)

    with open(save_file, 'rb') as f:
        return pickle.load(f)


if __name__ == "__main__":
    _init_pkl(read_file, save_file)
    dataset = load_beam_name(read_file, save_file)
    print(dataset.head())
    # print(dataset['樓層'])
    # print(list(dataset))
