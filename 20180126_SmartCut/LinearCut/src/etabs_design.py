""" load etbas design
"""
import math
import pandas as pd


def load_etabs_design(read_file):
    """ load etabs beam design
    """
    return pd.read_excel(
        read_file, sheet_name='Concrete_Design_2___Beam_Summar')


def post_e2k(df, e2k):
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

    grouped = df.groupby(['Story', 'BayID'])['StnLoc']

    df['LSupportWidth'] = round(grouped.transform('min'), 3)
    df['RSupportWidth'] = round(df['Length'] - grouped.transform('max'), 3)

    return df


def post_beam_name(df, beam_name):
    """ add beam/frame name id to etabs_design
    """
    df = df.assign(BeamID='', FrameID='')

    for (story, bay_id), group in df.groupby(['Story', 'BayID'], sort=False):
        beam_id, frame_id = beam_name.loc[(story, bay_id), :].values
        group = group.assign(BeamID=beam_id, FrameID=frame_id)
        df.loc[
            group.index, ['BeamID', 'FrameID']] = group[['BeamID', 'FrameID']]

    return df


def main():
    """
    test
    """
    from tests.const import const
    from src.e2k import load_e2k
    from src.beam_name import load_beam_name

    etabs_design = load_etabs_design(const['etabs_design_path'])
    print(etabs_design.head())

    e2k = load_e2k(const['e2k_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    print(etabs_design.head())

    beam_name = load_beam_name(const['beam_name_path'])
    etabs_design = post_beam_name(etabs_design, beam_name)
    print(etabs_design.head())


if __name__ == "__main__":
    main()
