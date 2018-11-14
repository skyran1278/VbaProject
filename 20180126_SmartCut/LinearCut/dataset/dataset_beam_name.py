import os
import sys
import pickle

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

read_file = SCRIPT_DIR + '/first_run.xlsx'
save_file = SCRIPT_DIR + '/beam_name.pkl'


def _load_name():
    return pd.read_excel(SCRIPT_DIR + '/first_run.xlsx', sheet_name='梁名編號', index_col=[0, 1], usecols=[1, 2, 3, 4])


def init_pkl():
    dataset = _load_name()

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_beam_name():
    if not os.path.exists(save_file):
        init_pkl()

    with open(save_file, 'rb') as f:
        return pickle.load(f)


if __name__ == "__main__":
    init_pkl()
    dataset = load_beam_name()
    print(dataset.head())
    # print(dataset['樓層'])
    # print(list(dataset))
