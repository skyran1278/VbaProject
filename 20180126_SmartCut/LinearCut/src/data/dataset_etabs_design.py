""" load etbas design
"""
import os
import pickle

import pandas as pd

# SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))


# read_file = f'{SCRIPT_DIR}/{BEAM_DESIGN}'
# save_file = f'{SCRIPT_DIR}/../temp/{BEAM_DESIGN}.pkl'


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
    with open(save_file, 'wb') as filename:
        pickle.dump(dataset, filename, True)
    print("Done!")


def load_beam_design(read_file, save_file):
    """ load etabs beam design
    """
    if not os.path.exists(save_file):
        _init_pkl(read_file, save_file)

    with open(save_file, 'rb') as filename:
        dataset = pickle.load(filename)

    return dataset


if __name__ == '__main__':
    from const import BEAM_DESIGN

    READ_FILE = f'{BEAM_DESIGN}'
    SAVE_FILE = f'{BEAM_DESIGN}.pkl'

    _init_pkl(READ_FILE, SAVE_FILE)
    DATASET = load_beam_design(READ_FILE, SAVE_FILE)
    print(DATASET.head())
    print(DATASET[['Story', 'VRebar']])
