import os
import pickle
import pandas as pd


def _create_pkl(save_file, dataset):
    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")

    return dataset


def load_pkl(save_file, dataset=None):
    if os.path.isfile(save_file):
        with open(save_file, 'rb') as f:
            dataset = pickle.load(f)
    else:
        dataset = _create_pkl(save_file, dataset)

    return dataset
