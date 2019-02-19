""" load etbas design
"""
import os
import pickle

import pandas as pd


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


def load_beam_design(read_file, save_file):
    """ load etabs beam design
    """
    if not os.path.exists(save_file):
        _init_pkl(read_file, save_file)

    with open(save_file, 'rb') as f:
        dataset = pickle.load(f)

    return dataset


if __name__ == '__main__':
    from const import const
    ETABS_DESIGN_PATH = const['etabs_design_path']
    READ_FILE = ETABS_DESIGN_PATH
    SAVE_FILE = f'{ETABS_DESIGN_PATH}.pkl'

    _init_pkl(READ_FILE, SAVE_FILE)
    DATASET = load_beam_design(READ_FILE, SAVE_FILE)
    print(DATASET.head())
    print(DATASET[['Story', 'VRebar']])
