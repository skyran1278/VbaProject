import os
import sys
import pytest

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.pkl import load_pkl

from database.dataset_beam_design import load_beam_design
from database.dataset_e2k import load_e2k
from tests.const import BEAM_DESIGN, E2K
from app.init_table import init_beam, init_beam_name
from app.stirrups import calc_sturrups
from app.bar_size_num import calc_db_by_beam, calc_db_by_frame
from app.bar_ld import calc_ld, add_ld
from app.bar_con import cut_conservative, add_simple_ld
from app.bar_cut import cut_optimization

# @pytest.fixture(scope='session')


@pytest.fixture(scope='session')
def beam_design():
    beam_design = load_beam_design(
        f'{SCRIPT_DIR}/{BEAM_DESIGN}', f'{SCRIPT_DIR}/../temp/{BEAM_DESIGN}.pkl')

    return beam_design


@pytest.fixture(scope='session')
def e2k():
    e2k = load_e2k(
        f'{SCRIPT_DIR}/{E2K}', f'{SCRIPT_DIR}/../temp/{E2K}.pkl')

    return e2k


def first_full_run():
    full_run(multi=3, calc_db=calc_db_by_beam, path='/../out/first_run.xlsx')


def second_run():
    full_run(multi=3, calc_db=calc_db_by_frame, path='/../out/second_run.xlsx')


def full_run(multi, calc_db, path):
    # 初始化輸出表格
    beam_name = init_beam_name()
    beam_con = init_beam()
    beam_cut = init_beam(multi)

    # 計算箍筋
    beam_con, beam_v = calc_sturrups(beam_con)
    beam_cut.loc[:, [('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右')]] = beam_con[[
        ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右')]]
    (beam_cut, beam_con, beam_v) = load_pkl(
        SCRIPT_DIR + '/stirrups.pkl', (beam_cut, beam_con, beam_v))

    # 以一根梁為單位 計算主筋 first run
    # 以一台梁為單位 計算主筋 second run
    beam_v_m = calc_db(beam_v)
    beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl', beam_v_m)

    # 計算延伸長度
    beam_v_m_ld = calc_ld(beam_v_m)

    # 加上延伸長度
    beam_ld_added = add_ld(beam_v_m_ld)
    # 傳統 端點加上簡算法的延伸長度
    beam_ld_added = add_simple_ld(beam_ld_added)
    beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl', beam_ld_added)
    # beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl')

    # 傳統斷筋
    beam_con = cut_conservative(beam_ld_added, beam_con)

    # 多點斷筋
    beam_cut = cut_optimization(multi, beam_ld_added, beam_cut)
