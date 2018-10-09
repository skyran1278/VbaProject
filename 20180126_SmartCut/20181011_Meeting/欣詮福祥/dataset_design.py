import os
import pickle
import numpy as np

dataset_dir = os.path.dirname(os.path.abspath(__file__))
save_file = dataset_dir + "/beam_design.pkl"


def _load_file(file_name):
    file_path = dataset_dir + "/" + file_name
    dataset = np.genfromtxt(
        file_path, dtype=None, names=True, delimiter='\t', encoding='utf8')

    return dataset


def init_pkl():
    dataset = _load_file(
        'Concrete Design 2 Beam Summary Data ACI 318-05 IBC 2003.txt')

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
    print(dataset[['Story', 'BayID']][682])


if __name__ == '__main__':
    main()
