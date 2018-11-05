import os
import pickle
import pandas as pd


def create_pkl(save_file, dataset):
    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")

    return dataset


def load_pkl(save_file, dataset=None):
    if not os.path.exists(save_file):
        dataset = create_pkl(save_file, dataset)

    with open(save_file, 'rb') as f:
        dataset = pickle.load(f)

    return dataset
