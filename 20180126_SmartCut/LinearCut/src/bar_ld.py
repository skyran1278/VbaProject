""" calculate ld
"""
import numpy as np

from src.dataset_rebar import double_area, rebar_area, rebar_db


def _double_size_area(real_v_size):
    rebar_num = real_v_size[0]

    return np.where(rebar_num == '2', double_area(real_v_size), rebar_area(real_v_size))


def _ld(df, loc, e2k, cover):
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
    # pylint: disable=invalid-name
    B = df['SecID'].apply(lambda x: sections[x, 'B']) * 100
    material = df['SecID'].apply(lambda x: sections[x, 'MATERIAL'])
    # pylint: disable=invalid-name
    fc = material.apply(lambda x: materials[x, 'FC']) / 10
    # pylint: disable=invalid-name
    fy = material.apply(lambda x: materials[x, 'FY']) / 10
    fyh = fy
    cover = cover * 100
    db = df[bar_size].apply(rebar_db) * 100  # pylint: disable=invalid-name
    num = df[bar_1st]
    dh = df['RealVSize'].apply(rebar_db) * 100  # pylint: disable=invalid-name
    avh = df['RealVSize'].apply(_double_size_area) * 10000
    spacing = df['RealSpacing'] * 100

    # 5.2.2
    fc[np.sqrt(fc) > 26.5] = 700

    # R5.3.4.1.1
    cc = dh + cover  # pylint: disable=invalid-name

    # R5.3.4.1.1
    cs = (B - db * num - dh * 2 - cover * 2) / \
        (num - 1) / 2  # pylint: disable=invalid-name

    # Vertical splitting failure / Horizontal splitting failure
    cb = np.where(cc <= cs, cc, cs) + db / 2  # pylint: disable=invalid-name

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
    ld = 0.28 * fy / np.sqrt(fc) * db / np.minimum((cb + ktr) /
                                                   db, 2.5)  # pylint: disable=invalid-name

    # 5.3.4.1
    simple_ld = 0.19 * fy / np.sqrt(fc) * db

    # phi_s factor
    ld[db < 2.2] = 0.8 * ld
    simple_ld[db < 2.2] = 0.8 * simple_ld

    # phi_t factor
    if loc == 'Top':
        ld = 1.3 * ld  # pylint: disable=invalid-name
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


def calc_ld(etbas_design, e2k, const):
    """
    It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
    Reference:土木401-93
    PI = 3.1415926
    """
    rebar, cover = const['rebar'], const['cover']

    for loc in rebar:
        etbas_design = etbas_design.assign(
            **_ld(etbas_design, loc, e2k, cover))

    return etbas_design


def add_ld(etbas_design, ld_type, rebar):
    """
    add ld
    ld_type: 'Ld', 'SimpleLd' I think 'SimpleLd' maybe not necessary
    """
    ld_design = etbas_design.copy()

    def init_ld(df):
        return {
            bar_num_ld: df[bar_num],
        }

    # 好像可以不用分上下層
    # 分比較方便
    for loc in rebar:
        bar_num = 'Bar' + loc + 'Num'
        ld = loc + ld_type  # pylint: disable=invalid-name
        bar_num_ld = bar_num + ld_type

        ld_design = ld_design.assign(**init_ld(ld_design))

        count = 0

        for name, group in ld_design.groupby(['Story', 'BayID'], sort=False):
            group = group.copy()

            if ld_type == 'Ld':
                iteration = range(len(group))
            elif ld_type == 'SimpleLd':
                iteration = (0, -1)

            for row in iteration:
                stn_loc = group.at[group.index[row], 'StnLoc']
                stn_ld = group.at[group.index[row], ld]
                stn_inter = (group['StnLoc'] >= stn_loc -
                             stn_ld) & (group['StnLoc'] <= stn_loc + stn_ld)
                group.loc[stn_inter, bar_num_ld] = np.maximum(
                    group.at[group.index[row], bar_num], group.loc[stn_inter, bar_num_ld])

            ld_design.loc[group.index, bar_num_ld] = group[bar_num_ld]

            count += 1
            if count % 100 == 0:
                print(name)

    return ld_design


def main():
    """
    test
    """
    from src.init_beam import init_beam
    from src.const import const
    from src.dataset_etabs_design import load_beam_design
    from src.dataset_e2k import load_e2k
    from src.execution_time import Execution
    from src.stirrups import calc_stirrups
    from src.bar_size_num import calc_db

    e2k_path, etabs_design_path = const['e2k_path'], const['etabs_design_path']

    e2k = load_e2k(e2k_path, e2k_path + '.pkl')
    etabs_design = load_beam_design(
        etabs_design_path, etabs_design_path + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3)
    execution = Execution()
    beam, dh_design = calc_stirrups(
        beam, etabs_design, e2k, const, consider_vc=False)

    db_design = calc_db('BayID', dh_design, e2k, const)

    execution.time('ld')
    ld_design = calc_ld(db_design, e2k, const)
    print(ld_design.head())
    execution.time('ld')

    execution.time('add_ld')
    ld_design = add_ld(ld_design, 'Ld', const['rebar'])
    print(ld_design.head())
    execution.time('add_ld')

    execution.time('add_simple_ld')
    ld_design = add_ld(ld_design, 'SimpleLd', const['rebar'])
    print(ld_design.head())
    execution.time('add_simple_ld')


if __name__ == "__main__":
    main()
