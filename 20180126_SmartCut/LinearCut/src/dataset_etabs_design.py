""" load etbas design
"""
import pandas as pd


def load_beam_design(read_file):
    """ load etabs beam design
    """
    return pd.read_excel(
        read_file, sheet_name='Concrete_Design_2___Beam_Summar')


def merge_e2k_to_etbas_design(df, e2k):
    pass


def main():
    """
    test
    """
    from tests.const import const

    dataset = load_beam_design(const['etabs_design_path'])
    print(dataset.head())


if __name__ == "__main__":
    main()
