""" pkl functions """
import pickle


def init_pkl(save_file, dataset):
    """
    init pkl
    """
    print("Creating pickle file ...")
    with open(save_file, 'wb') as filename:
        pickle.dump(dataset, filename, True)
    print("Done!")

    # return dataset


def load_pkl(save_file):
    """ if no dataset, create pkl, or load pkl """
    # if dataset is not None:
    #     dataset = _init_pkl(dataset, save_file)
    # else:
    with open(save_file, 'rb') as filename:
        dataset = pickle.load(filename)

    # if not os.path.exists(save_file):
    #     dataset = _create_pkl(save_file, dataset)

    return dataset
