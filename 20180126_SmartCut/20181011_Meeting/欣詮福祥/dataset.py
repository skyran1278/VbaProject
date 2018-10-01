import os
import pickle
import numpy as np

dataset_dir = os.path.dirname(os.path.abspath(__file__))
save_file = dataset_dir + "/moment_diagram.pkl"


def _load_file(file_name, usecols=None):
    file_path = dataset_dir + "/" + file_name
    beam_forces = np.genfromtxt(
        file_path, dtype=None, names=True, delimiter='\t', encoding='utf8', usecols=usecols)

    return beam_forces


def init_moment():
    dataset = {}

    beam_forces = _load_file('Beam Forces.txt')
    frame_assignments_summary = _load_file(
        'Frame Assignments Summary.txt', usecols=(0, 1, 2, 3))

    dataset['beam_force'] = beam_forces
    dataset['beam_length'] = frame_assignments_summary

    print("Creating pickle file ...")
    with open(save_file, 'wb') as f:
        pickle.dump(dataset, f, True)
    print("Done!")


def load_moment():
    if not os.path.exists(save_file):
        init_moment()

    with open(save_file, 'rb') as f:
        dataset = pickle.load(f)

    return dataset['beam_force'], dataset['beam_length']


def main():
    b_force, b_length = load_moment()
    print(b_force[1]['Story'])
    print(b_length['Length'][0])


if __name__ == '__main__':
    main()
