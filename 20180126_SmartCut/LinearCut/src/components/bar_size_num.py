""" calculate rebar size and num by moment
"""
import numpy as np

# SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

# from utils.pkl import load_pkl

from const import BAR, DB_SPACING, COVER

# from data.dataset_etabs_design import load_beam_design
# from data.dataset_e2k import load_e2k
# from data.dataset_beam_name import load_beam_name
from data.dataset_rebar import rebar_db, rebar_area

# save_file = SCRIPT_DIR + '/3pionts.xlsx'
# stirrups_save_file = SCRIPT_DIR + '/stirrups.pkl'

# e2k = load_e2k()
# sections = e2k['sections']
# etabs_design = load_beam_design()
# beam_3p = init_beam_3points_table()
# beam_3p, beam_design_table_stirrups = calc_sturrups(beam_3p)
# beam_3points_table = init_beam_3points_table()
# beam_3points_table, beam_design_table_stirrups = calc_sturrups(
#     beam_3points_table)
# (beam_3points_table, beam_design_table_stirrups) = load_pkl(
#     stirrups_save_file, (beam_3points_table, beam_design_table_stirrups))
# (beam_3p, etabs_design) = load_pkl(stirrups_save_file)


def _bar_name(loc):
    bar_size = 'Bar' + loc + 'Size'
    bar_num = 'Bar' + loc + 'Num'
    bar_cap = 'Bar' + loc + 'Cap'
    bar_1st = 'Bar' + loc + '1st'
    bar_2nd = 'Bar' + loc + '2nd'

    return (bar_size, bar_num, bar_cap, bar_1st, bar_2nd)


def _calc_bar_size_num(i, loc, e2k):
    sections = e2k['sections']
    bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(loc)

    def _calc_capacity(df):
        # dh = df['VSize'].apply()
        # 應該可以用 apply 來改良，晚點再來做
        # 這裡應該拿最後配的來算，但是因為號數整支梁都會相同，所以沒差
        # 後來查了一下 發現好像差不多
        dh = df['VSize'].apply(lambda x: rebar_db('#' + x.split('#')[1]))
        # dh = np.array([rebar_db('#' + v_size.split('#')[1])
        #    for v_size in df['VSize']])
        db = rebar_db(BAR[loc][i])
        width = df['SecID'].apply(lambda x: sections[(x, 'B')])
        # width = np.array([sections[(sec_ID, 'B')] for sec_ID in df['SecID']])
        # cover = np.array([sections[(sec_ID, 'COVER' + loc)] for sec_ID in df['SecID']])

        return np.floor((width - 2 * COVER - 2 * dh - db) / (DB_SPACING * db + db)) + 1
        # return np.ceil((width - 2 * COVER - 2 * dh - db) / (DB_SPACING * db + db))

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
        # for i in range(len(df.index)):
        #     if num:
        #         pass
        # df[bar_num] - df[bar_cap] == 1
        #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(cap_num - 1)
        #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = concat_size(2)
        # elif loc_num > cap_num:
        #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(cap_num)
        #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = concat_size(loc_num - cap_num)
        # else:
        #     beam_3p.loc[i, ('主筋', bar_loc)] = concat_size(loc_num)
        #     beam_3p.loc[i + to_2nd, ('主筋', bar_loc)] = 0

    return {
        bar_size: BAR[loc][i],
        bar_cap: _calc_capacity,
        # 增加扣 0.05 的容量
        bar_num: lambda x: np.maximum(np.ceil(x['As' + loc] / rebar_area(BAR[loc][i]) - 0.05), 2),
        bar_1st: _calc_1st,
        bar_2nd: _calc_2nd
    }


def calc_db(by, etabs_design, e2k):
    """ calculate db by beam or usr defined frame, should first calculate stirrups
    """
    db_design = etabs_design.copy()

    for loc in BAR:

        bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(loc)

        db_design = db_design.assign(
            **_calc_bar_size_num(0, loc, e2k))

        for _, group in db_design.groupby(['Story', by], sort=False):
            i = 0

            while np.any(group[bar_num] > 2 * group[bar_cap]):
                i += 1
                group = group.assign(**_calc_bar_size_num(i, loc, e2k))

            db_design.loc[group.index.tolist(), [bar_size, bar_num, bar_cap, bar_1st, bar_2nd]
                          ] = group[[bar_size, bar_num, bar_cap, bar_1st, bar_2nd]]

    return db_design


# def calc_db_by_frame(etabs_design, e2k):
#     """ calculate db by usr defined frame, should first calculate stirrups
#     """
#     db_design = etabs_design.copy()

#     for loc in BAR:
#         bar_size, bar_num, bar_cap, bar_1st, bar_2nd = _bar_name(loc)

#         i = 0

#         db_design = db_design.assign(
#             **_calc_bar_size_num(loc, i, e2k))

#         for _, group in db_design.groupby(['Story', 'FrameID'], sort=False):
#             i = 0

#             while np.any(group[bar_num] > 2 * group[bar_cap]):
#                 i += 1
#                 group = group.assign(**_calc_bar_size_num(loc, i, e2k))

#             db_design.loc[group.index, [bar_size, bar_num, bar_cap, bar_1st, bar_2nd]
#                           ] = group[[bar_size, bar_num, bar_cap, bar_1st, bar_2nd]]

#     return db_design


# def _add_beam_name(etabs_design):
#     """ calculate db by usr defined frame, should first calculate stirrups
#     """
    # beam_name = load_beam_name()
#     etabs_design = etabs_design.assign(BeamID='', FrameID='')

#     for (story, bayID), group in etabs_design.groupby(['Story', 'BayID'], sort=False):
#         beamID, frameID = beam_name.loc[(story, bayID), :]
#         group = group.assign(BeamID=beamID, FrameID=frameID)
#         etabs_design.loc[group.index, ['BeamID', 'FrameID']
#                          ] = group[['BeamID', 'FrameID']]

#     return etabs_design

def main():
    """ test
    """
    from components.init_beam import init_beam, add_and_alter_beam_id
    from const import E2K_PATH, ETABS_DESIGN_PATH, BEAM_NAME_PATH
    from data.dataset_etabs_design import load_beam_design
    from data.dataset_e2k import load_e2k
    from data.dataset_beam_name import load_beam_name
    from utils.execution_time import Execution
    from components.stirrups import calc_stirrups

    e2k = load_e2k(E2K_PATH, E2K_PATH + '.pkl')
    etabs_design = load_beam_design(
        ETABS_DESIGN_PATH, ETABS_DESIGN_PATH + '.pkl')
    beam_name = load_beam_name(BEAM_NAME_PATH, BEAM_NAME_PATH + '.pkl')

    beam = init_beam(etabs_design, e2k, moment=3, shear=True)
    execution = Execution()
    beam, dh_design = calc_stirrups(beam, etabs_design)

    # (_, etabs_design) = load_pkl(stirrups_save_file)
    execution.time('BayID')
    db_design = calc_db('BayID', dh_design, e2k)
    print(db_design.head())
    execution.time('BayID')

    execution.time('FrameID')
    beam, dh_design = add_and_alter_beam_id(
        beam, beam_name, dh_design)
    db_design = calc_db('FrameID', dh_design, e2k)
    print(db_design.head())
    execution.time('FrameID')
    # db_design = load_pkl(SCRIPT_DIR + '/db_design.pkl', db_design)


if __name__ == "__main__":
    main()
