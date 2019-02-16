""" calculate ld
"""
import numpy as np

from const import BAR, COVER
from data.dataset_rebar import double_area, rebar_area, rebar_db


def _double_size_area(real_v_size):
    rebar_num = real_v_size[0]

    return np.where(rebar_num == '2', double_area(real_v_size), rebar_area(real_v_size))


def _ld(df, loc, e2k):
    """
    It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
    Reference:土木401-93
    PI = 3.1415926
    """
    materials, sections = e2k['materials'], e2k['sections']

    bar_size = 'Bar' + loc + 'Size'
    bar_1st = 'Bar' + loc + '1st'

    # 延伸長度比較熟悉 cm 操作
    # m => cm
    B = df['SecID'].apply(lambda x: sections[x, 'B']) * 100
    material = df['SecID'].apply(lambda x: sections[x, 'MATERIAL'])
    fc = material.apply(lambda x: materials[x, 'FC']) / 10
    fy = material.apply(lambda x: materials[x, 'FY']) / 10
    fyh = fy
    cover = COVER * 100
    db = df[bar_size].apply(rebar_db) * 100
    num = df[bar_1st]
    dh = df['RealVSize'].apply(rebar_db) * 100
    avh = df['RealVSize'].apply(_double_size_area) * 10000
    spacing = df['RealSpacing'] * 100

    # 5.2.2
    fc[np.sqrt(fc) > 26.5] = 700

    # R5.3.4.1.1
    cc = dh + cover

    # R5.3.4.1.1
    cs = (B - db * num - dh * 2 - cover * 2) / (num - 1) / 2

    # Vertical splitting failure / Horizontal splitting failure
    cb = np.where(cc <= cs, cc, cs) + db / 2

    # R5.3.4.1.2
    ktr = np.where(cc <= cs, 1, 2 / num) * avh * fyh / 105 / spacing

    # if cs > cc:
    #     # Vertical splitting failure
    #     cb = db / 2 + cc
    #     # R5.3.4.1.2
    #     ktr = (PI * dh ** 2 / 4) * fyh / 105 / spacing
    # else:
    #     # Horizontal splitting failure
    #     cb = db / 2 + cs
    #     # R5.3.4.1.2
    #     ktr = 2 * (PI * dh ** 2 / 4) * fyh / 105 / spacing / num

    # 5.3.4.1
    ld = 0.28 * fy / np.sqrt(fc) * db / np.minimum((cb + ktr) / db, 2.5)

    # 5.3.4.1
    simple_ld = 0.19 * fy / np.sqrt(fc) * db

    # phi_s factor
    ld[db < 2.2] = 0.8 * ld
    simple_ld[db < 2.2] = 0.8 * simple_ld

    # phi_t factor
    if loc == 'Top':
        ld = 1.3 * ld
        simple_ld = 1.3 * simple_ld

    ld[ld > simple_ld] = simple_ld

    # 5.3.1
    ld[ld < 30] = 30
    simple_ld[simple_ld < 30] = 30

    return {
        # cm => m
        loc + 'Ld': ld / 100,
        loc + 'SimpleLd': simple_ld / 100
    }


def calc_ld(etbas_design, e2k):
    """
    It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
    Reference:土木401-93
    PI = 3.1415926
    """

    for loc in BAR:
        etbas_design = etbas_design.assign(**_ld(etbas_design, loc, e2k))

    return etbas_design


def add_ld(etbas_design):
    """
    add ld
    """
    ld_design = etbas_design.copy()

    def init_ld(df):
        return {
            bar_num_ld: df[bar_num],
        }

    # 好像可以不用分上下層
    # 分比較方便
    for loc in BAR:
        bar_num = 'Bar' + loc + 'Num'
        ld = loc + 'Ld'
        bar_num_ld = bar_num + 'Ld'

        ld_design = ld_design.assign(**init_ld(ld_design))

        count = 0

        for name, group in ld_design.groupby(['Story', 'BayID'], sort=False):
            group = group.copy()
            for i in range(len(group)):
                stn_loc = group.at[group.index[i], 'StnLoc']
                stn_ld = group.at[group.index[i], ld]
                stn_inter = (group['StnLoc'] >= stn_loc -
                             stn_ld) & (group['StnLoc'] <= stn_loc + stn_ld)
                group.loc[stn_inter, bar_num_ld] = np.maximum(
                    group.at[group.index[i], bar_num], group.loc[stn_inter, bar_num_ld])

            ld_design.loc[group.index, bar_num_ld] = group[bar_num_ld]
            count += 1
            if count % 100 == 0:
                print(name)

    return ld_design


def main():
    """
    test
    """
    from components.init_beam import init_beam
    from const import E2K_PATH, ETABS_DESIGN_PATH
    from data.dataset_etabs_design import load_beam_design
    from data.dataset_e2k import load_e2k
    from utils.execution_time import Execution
    from components.stirrups import calc_stirrups
    from components.bar_size_num import calc_db

    e2k = load_e2k(E2K_PATH, E2K_PATH + '.pkl')
    etabs_design = load_beam_design(
        ETABS_DESIGN_PATH, ETABS_DESIGN_PATH + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3, shear=True)
    execution = Execution()
    beam, dh_design = calc_stirrups(beam, etabs_design)

    db_design = calc_db('BayID', dh_design, e2k)

    execution.time('ld')
    ld_design = calc_ld(db_design, e2k)
    print(ld_design.head())
    execution.time('ld')

    execution.time('add_ld')
    ld_design = add_ld(ld_design)
    print(ld_design.head())
    execution.time('add_ld')


if __name__ == "__main__":
    main()
