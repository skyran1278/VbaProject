""" calculate ld
"""
import numpy as np

from src.rebar import double_area, rebar_area, rebar_db


def _double_size_area(real_v_size):
    rebar_num = real_v_size[0]

    return np.where(rebar_num == '2', double_area(real_v_size), rebar_area(real_v_size))


def _ld(df, loc, cover):
    """
    It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
    Reference:土木401-93
    PI = 3.1415926
    """
    bar_size = 'Bar' + loc + 'Size'
    bar_1st = 'Bar' + loc + '1st'

    # 延伸長度比較熟悉 cm 操作
    # m => cm
    # pylint: disable=invalid-name
    B = df['B'] * 100
    # pylint: disable=invalid-name
    fc = df['Fc'] / 10
    # pylint: disable=invalid-name
    fy = df['Fy'] / 10
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


def calc_ld(etbas_design, const):
    """
    It is used for nominal concrete in case of phi_e=1.0 & phi_t=1.0.
    Reference:土木401-93
    PI = 3.1415926
    """
    rebar, cover = const['rebar'], const['cover']

    for loc in rebar:
        etbas_design = etbas_design.assign(
            **_ld(etbas_design, loc, cover))

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
                stn_inter = (
                    (group['StnLoc'] >= stn_loc - stn_ld) &
                    (group['StnLoc'] <= stn_loc + stn_ld)
                )
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
    from src.execution_time import Execution
    from tests.const import const
    from src.beam import init_beam
    from src.e2k import load_e2k
    from src.etabs_design import load_etabs_design, post_e2k
    from src.stirrups import calc_stirrups
    from src.bar_size_num import calc_db

    execution = Execution()

    e2k = load_e2k(const['e2k_path'])
    etabs_design = load_etabs_design(const['etabs_design_path'])
    etabs_design = post_e2k(etabs_design, e2k)
    beam = init_beam(etabs_design)
    beam, etabs_design = calc_stirrups(beam, etabs_design, const)
    etabs_design = calc_db('BayID', etabs_design, const)

    execution.time('ld')
    etabs_design = calc_ld(etabs_design, const)
    print(etabs_design.head())
    execution.time('ld')

    execution.time('add_ld')
    etabs_design = add_ld(etabs_design, 'Ld', const['rebar'])
    print(etabs_design.head())
    execution.time('add_ld')

    # 非常有可能會搞錯，不是所有都加 simple ld，只有兩端，所以建議可以不要這個。
    execution.time('add_simple_ld')
    etabs_design = add_ld(etabs_design, 'SimpleLd', const['rebar'])
    print(etabs_design.head(100))
    execution.time('add_simple_ld')


if __name__ == "__main__":
    main()
