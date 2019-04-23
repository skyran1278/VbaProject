""" load user defined beam name
"""
import pandas as pd


def load_beam_name(read_file):
    """ load beam name
    """

    return pd.read_excel(read_file, sheet_name='梁名編號', index_col=[0, 1], usecols=[1, 2, 3, 4])


def main():
    """
    test
    """
    from tests.const import const

    dataset = load_beam_name(const['beam_name_path'])
    print(dataset.head())


if __name__ == "__main__":
    main()
