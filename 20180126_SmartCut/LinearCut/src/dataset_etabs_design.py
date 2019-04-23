""" load etbas design
"""
import math
import pandas as pd


def load_beam_design(read_file):
    """ load etabs beam design
    """
    return pd.read_excel(
        read_file, sheet_name='Concrete_Design_2___Beam_Summar')


def merge_e2k_to_etbas_design(df, e2k):
    """
    merge e2k imformation to etabs design
    """
    coors = e2k['point_coordinates']
    lines = e2k['lines']
    materials = e2k['materials']
    sections = e2k['sections']

    section_material = df['SecID'].apply(
        lambda x: sections[x, 'MATERIAL']
    )

    line_id = df['BayID'].apply(lambda x: lines[x, 'BEAM'])

    df['B'] = df['SecID'].apply(lambda x: sections[x, 'B'])
    df['Fc'] = section_material.apply(lambda x: materials[x, 'FC'])
    df['Fy'] = section_material.apply(lambda x: materials[x, 'FY'])
    df['Fy'] = line_id.apply(
        lambda x: math.sqrt(sum((coors[x[1]] - coors[x[0]]) ** 2)))

    df.groupby('BayID').transform()
    df.groupby(['Story', 'BayID']).transform()


def main():
    """
    test
    """
    from tests.const import const
    from src.dataset_e2k import load_e2k

    e2k = load_e2k(const['e2k_path'])
    dataset = load_beam_design(const['etabs_design_path'])
    merge_e2k_to_etbas_design(dataset, e2k)
    print(dataset.head())


if __name__ == "__main__":
    main()
