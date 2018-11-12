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
    if dataset is not None:
        dataset = _create_pkl(save_file, dataset)
    else:
        with open(save_file, 'rb') as f:
            dataset = pickle.load(f)

    # if not os.path.exists(save_file):
    #     dataset = _create_pkl(save_file, dataset)

    return dataset
