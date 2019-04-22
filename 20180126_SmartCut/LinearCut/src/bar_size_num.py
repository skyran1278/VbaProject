""" calculate rebar size and num by moment
"""
import numpy as np

from src.dataset_rebar import rebar_db, rebar_area


def _bar_name(loc):
    bar_size = 'Bar' + loc + 'Size'
    bar_num = 'Bar' + loc + 'Num'
    bar_cap = 'Bar' + loc + 'Cap'
    bar_1st = 'Bar' + loc + '1st'
    bar_2nd = 'Bar' + loc + '2nd'

    return (bar_size, bar_num, bar_cap, bar_1st, bar_2nd)


def _calc_bar_size_num(rebar_i, loc, e2k, const):
    sections = e2k['sections']
    rebar, db_spacing, cover = const['rebar'], const['db_spacing'], const['cover']

    bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(loc)

    def _calc_capacity(df):
        # dh = df['VSize'].apply()
        # 應該可以用 apply 來改良，晚點再來做
        # 這裡應該拿最後配的來算，但是因為號數整支梁都會相同，所以沒差
        # 後來查了一下 發現好像差不多
        # pylint: disable=invalid-name
        dh = df['VSize'].apply(lambda x: rebar_db('#' + x.split('#')[1]))
        # pylint: disable=invalid-name
        db = rebar_db(rebar[loc][rebar_i])
        width = df['SecID'].apply(lambda x: sections[(x, 'B')])

        return np.floor((width - 2 * cover - 2 * dh - db) / (db_spacing * db + db)) + 1

    def _calc_1st(df):
        bar_1st = np.where(df[bar_num] > df[bar_cap],
                           df[bar_cap], df[bar_num])
        bar_1st[df[bar_num] - df[bar_cap] ==
                1] = df[bar_cap][df[bar_num] - df[bar_cap] == 1] - 1

        return bar_1st

    def _calc_2nd(df):
        bar_2nd = np.where(df[bar_num] > df[bar_cap],
                           df[bar_num] - df[bar_cap], 0)
        bar_2nd[df[bar_num] - df[bar_cap] == 1] = 2

        return bar_2nd

    return {
        bar_size: rebar[loc][rebar_i],
        bar_cap: _calc_capacity,
        # 增加扣 0.05 的容量
        bar_num: lambda x: np.maximum(
            np.ceil(x['As' + loc] / rebar_area(rebar[loc][rebar_i]) - 0.05), 2
        ),
        bar_1st: _calc_1st,
        bar_2nd: _calc_2nd
    }


def calc_db(by, etabs_design, e2k, const):  # pylint: disable=invalid-name
    """ calculate db by beam or usr defined frame, should first calculate stirrups
    """
    rebar = const['rebar']

    db_design = etabs_design.copy()

    for loc in rebar:
        bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(loc)

        db_design = db_design.assign(
            **_calc_bar_size_num(0, loc, e2k, const))

        for _, group in db_design.groupby(['Story', by], sort=False):
            rebar_i = 0

            while np.any(group[bar_num] > 2 * group[bar_cap]):
                rebar_i += 1
                group = group.assign(
                    **_calc_bar_size_num(rebar_i, loc, e2k, const))

            db_design.loc[group.index, [bar_size, bar_num, bar_cap, bar_1st, bar_2nd]
                          ] = group[[bar_size, bar_num, bar_cap, bar_1st, bar_2nd]]

    return db_design


def main():
    """ test
    """
    from src.init_beam import init_beam, add_and_alter_beam_id
    from src.const import const

    from src.dataset_etabs_design import load_beam_design
    from src.dataset_e2k import load_e2k
    from src.dataset_beam_name import load_beam_name
    from src.execution_time import Execution
    from src.stirrups import calc_stirrups

    e2k_path, etabs_design_path, beam_name_path = const[
        'e2k_path'], const['etabs_design_path'], const['beam_name_path']

    e2k = load_e2k(e2k_path, e2k_path + '.pkl')
    etabs_design = load_beam_design(
        etabs_design_path, etabs_design_path + '.pkl')
    beam_name = load_beam_name(beam_name_path, beam_name_path + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3)
    execution = Execution()
    beam, dh_design = calc_stirrups(beam, etabs_design, const)

    execution.time('BayID')
    db_design = calc_db('BayID', dh_design, e2k, const)
    print(db_design.head())
    execution.time('BayID')

    execution.time('FrameID')
    beam, dh_design = add_and_alter_beam_id(
        beam, beam_name, dh_design)
    db_design = calc_db('FrameID', dh_design, e2k, const)
    print(db_design.head())
    execution.time('FrameID')
    # db_design = load_pkl(SCRIPT_DIR + '/db_design.pkl', db_design)


if __name__ == "__main__":
    main()