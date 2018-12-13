import os
import sys
import pickle

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from database.const import BEAM_DESIGN

read_file = f'{SCRIPT_DIR}/{BEAM_DESIGN}'
save_file = f'{SCRIPT_DIR}/../temp/{BEAM_DESIGN}.pkl'


def _load_file(read_file):
    dataset = pd.read_excel(
        read_file, sheet_name='Concrete_Design_2___Beam_Summar')
    # dataset = np.genfromtxt(
    #     file_path, dtype=None, names=True, delimiter='\t', encoding='utf8')

    # dataset = pd.DataFrame(dataset)
    # print(dataset.head())
    return dataset


def _init_pkl(read_file, save_file):
    dataset = _load_file(read_file)

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_beam_design(read_file=read_file, save_file=save_file):
    if not os.path.exists(save_file):
        _init_pkl(read_file, save_file)

    with open(save_file, 'rb') as f:
        dataset = pickle.load(f)

    return dataset


if __name__ == '__main__':
    _init_pkl(read_file, save_file)
    dataset = load_beam_design(read_file, save_file)
    print(dataset.head())
    print(dataset[['Story', 'VRebar']])
