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
    def _cal_length(x):  # pylint: disable=invalid-name
        # x 距離 和 y 距離
        x_y_distance = coors[x[1]] - coors[x[0]]
        # 平方相加開根號
        length = math.sqrt(sum(x_y_distance ** 2))
        # 四捨五入到小數點下第三位
        return round(length, 3)

    def _min(x):  # pylint: disable=invalid-name
        return round(min(x), 3)

    def _max(x):  # pylint: disable=invalid-name
        return round(max(x), 3)

    coors = e2k['point_coordinates']
    lines = e2k['lines']
    mats = e2k['materials']
    secs = e2k['sections']

    sec_mats = df['SecID'].apply(lambda x: secs[x, 'MATERIAL'])
    line_ids = df['BayID'].apply(lambda x: lines[x, 'BEAM'])

    df['B'] = df['SecID'].apply(lambda x: secs[x, 'B'])
    df['H'] = df['SecID'].apply(lambda x: secs[x, 'H'])
    df['Fc'] = sec_mats.apply(lambda x: mats[x, 'FC'])
    df['Fy'] = sec_mats.apply(lambda x: mats[x, 'FY'])
    df['Length'] = line_ids.apply(_cal_length)

    df['LSupportWidth'] = df.groupby('BayID')['StnLoc'].transform(_min)
    df['RSupportWidth'] = (
        df['Length'] - df.groupby('BayID')['StnLoc'].transform(_max))

    return df


def main():
    """
    test
    """
    from tests.const import const
    from src.dataset_e2k import load_e2k

    e2k = load_e2k(const['e2k_path'])
    dataset = load_beam_design(const['etabs_design_path'])
    dataset = merge_e2k_to_etbas_design(dataset, e2k)
    print(dataset)


if __name__ == "__main__":
    main()
