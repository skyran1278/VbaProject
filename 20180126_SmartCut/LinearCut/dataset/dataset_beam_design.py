import os
import pickle
import pandas as pd
import numpy as np

from .const import BEAM_DESIGN

dataset_dir = os.path.dirname(os.path.abspath(__file__))

read_file = dataset_dir + '/' + BEAM_DESIGN + '.txt'
save_file = dataset_dir + '/' + BEAM_DESIGN + '.pkl'


def _load_file():
    dataset = pd.read_table(read_file, sep='\t')
    # dataset = np.genfromtxt(
    #     file_path, dtype=None, names=True, delimiter='\t', encoding='utf8')

    # dataset = pd.DataFrame(dataset)
    # print(dataset.head())
    return dataset


def init_pkl():
    dataset = _load_file()

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_beam_design():
    if not os.path.exists(save_file):
        init_pkl()

    with open(save_file, 'rb') as f:
        dataset = pickle.load(f)

    return dataset


def main():
    init_pkl()
    dataset = load_beam_design()
    print(dataset.head())
    print(dataset[['Story', 'VRebar']])


if __name__ == '__main__':
    main()
