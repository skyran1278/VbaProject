import os
import pickle

import pandas as pd
import numpy as np

dataset_dir = os.path.dirname(os.path.abspath(__file__))
read_file = dataset_dir + '/first_run.xlsx'
save_file = dataset_dir + '/beam_name.pkl'


def _load_name():
    return pd.read_excel(dataset_dir + '/first_run.xlsx', sheet_name='梁名編號', index_col=[0, 1], usecols=[1, 2, 3, 4])


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
