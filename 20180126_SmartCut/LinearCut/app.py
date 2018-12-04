import os
import sys

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.pkl import load_pkl
from utils.Clock import Clock
from init_table import init_beam, init_beam_name
from stirrups import calc_sturrups
from bar_size_num import calc_db_by_beam, calc_db_by_frame
from bar_ld import calc_ld, add_ld
from bar_con import cut_conservative, add_simple_ld
from bar_cut import cut_optimization


# 不管是物件導向設計還是函數式編程 只要能解決問題的就是好方法
# 現在還只是看的不爽 所以並沒有造成問題
# 物件導向是對於真實世界的物體的映射
# 函數式編程是對於資料更好的操控

clock = Clock()


def first_run():
    clock.time('梁名編號')

    writer = pd.ExcelWriter(SCRIPT_DIR + '/dataset/first_run.xlsx')
    beam_name = init_beam_name()
    beam_name.to_excel(writer, '梁名編號')
    writer.save()

    clock.time()


def full_run(multi, calc_db, path):
    writer = pd.ExcelWriter(SCRIPT_DIR + path)

    # 初始化輸出表格
    clock.time('初始化輸出表格')
    beam_name = init_beam_name()
    beam_con = init_beam()
    beam_cut = init_beam(multi)
    clock.time()

    # 計算箍筋
    clock.time('計算箍筋')
    beam_con, beam_v = calc_sturrups(beam_con)
    beam_cut.loc[:, [('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右')]] = beam_con[[
        ('箍筋', '左'), ('箍筋', '中'), ('箍筋', '右')]]
    (beam_cut, beam_con, beam_v) = load_pkl(
        SCRIPT_DIR + '/stirrups.pkl', (beam_cut, beam_con, beam_v))
    clock.time()

    # 以一根梁為單位 計算主筋 first run
    # 以一台梁為單位 計算主筋 second run
    clock.time('以一根梁為單位 計算主筋')
    beam_v_m = calc_db(beam_v)
    beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl', beam_v_m)
    clock.time()

    # 計算延伸長度
    clock.time('計算延伸長度')
    beam_v_m_ld = calc_ld(beam_v_m)
    clock.time()

    # 加上延伸長度
    clock.time('加上延伸長度')
    beam_ld_added = add_ld(beam_v_m_ld)
    # 傳統 端點加上簡算法的延伸長度
    beam_ld_added = add_simple_ld(beam_ld_added)
    beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl', beam_ld_added)
    # beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl')
    clock.time()

    # 傳統斷筋
    clock.time('傳統斷筋')
    beam_con = cut_conservative(beam_ld_added, beam_con)
    clock.time()

    # 三點斷筋
    clock.time('三點斷筋')
    beam_cut = cut_optimization(beam_ld_added, beam_cut)
    clock.time()

    # 輸出成表格
    clock.time('輸出成表格')
    beam_cut.to_excel(writer, '三點斷筋')
    beam_con.to_excel(writer, '傳統斷筋')
    beam_ld_added.to_excel(writer, 'beam_ld_added')
    beam_name.to_excel(writer, '梁名編號')
    writer.save()
    clock.time()


def first_full_run():
    full_run(multi=3, calc_db=calc_db_by_beam, path='/dataset/first_run.xlsx')
#     writer = pd.ExcelWriter(SCRIPT_DIR + '/dataset/first_run.xlsx')

#     # 初始化輸出表格
#     clock.time('初始化輸出表格')
#     beam_name = init_beam_name()
#     beam_3 = init_beam()
#     if multi == 5:
#         beam_5 = init_beam()
#     clock.time()
# # df2[['date', 'hour']] = df1[['date', 'hour']]
#     # 計算箍筋
#     clock.time('計算箍筋')
#     beam_3, beam_v = calc_sturrups(beam_3)
#     (beam_3, beam_v) = load_pkl(SCRIPT_DIR + '/stirrups.pkl', (beam_3, beam_v))
#     # (beam_3, beam_v) = load_pkl(SCRIPT_DIR + '/stirrups.pkl')
#     clock.time()

#     # 以一根梁為單位 計算主筋
#     clock.time('以一根梁為單位 計算主筋')
#     beam_v_m = calc_db_by_beam(beam_v)
#     beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl', beam_v_m)
#     clock.time()

#     # 計算延伸長度
#     clock.time('計算延伸長度')
#     beam_v_m_ld = calc_ld(beam_v_m)
#     clock.time()

#     # 加上延伸長度
#     clock.time('加上延伸長度')
#     beam_ld_added = add_ld(beam_v_m_ld)
#     # 傳統 端點加上簡算法的延伸長度
#     beam_ld_added = add_simple_ld(beam_ld_added)
#     beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl', beam_ld_added)
#     # beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl')
#     clock.time()

#     # 傳統斷筋
#     clock.time('傳統斷筋')
#     beam_3p_bar = cut_conservative(beam_ld_added, beam_3)
#     clock.time()

#     # 三點斷筋
#     clock.time('三點斷筋')
#     beam_3 = cut_optimization(beam_ld_added, beam_3)
#     clock.time()

#     # 輸出成表格
#     clock.time('輸出成表格')
#     beam_name.to_excel(writer, '梁名編號')
#     beam_3.to_excel(writer, '三點斷筋')
#     beam_3p_bar.to_excel(writer, '傳統斷筋')
#     beam_ld_added.to_excel(writer, 'beam_ld_added')
#     writer.save()
#     clock.time()


def second_run():
    full_run(multi=3, calc_db=calc_db_by_frame, path='/second_run.xlsx')
    # # second run
    # writer = pd.ExcelWriter(SCRIPT_DIR + '/second_run.xlsx')

    # # 初始化輸出表格
    # clock.time('初始化輸出表格')
    # beam_3p = init_beam()
    # clock.time()

    # # 計算箍筋
    # clock.time('計算箍筋')
    # beam_3p, beam_v = calc_sturrups(beam_3p)
    # (beam_3p, beam_v) = load_pkl(SCRIPT_DIR + '/stirrups.pkl', (beam_3p, beam_v))
    # clock.time()

    # # 以一台梁為單位 計算主筋 second run
    # clock.time('以一台梁為單位 計算主筋')
    # beam_v_m = calc_db_by_frame(beam_v)
    # beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl', beam_v_m)
    # clock.time()

    # # 計算延伸長度
    # clock.time('計算延伸長度')
    # beam_v_m_ld = calc_ld(beam_v_m)
    # clock.time()

    # # 加上延伸長度
    # clock.time('加上延伸長度')
    # beam_ld_added = add_ld(beam_v_m_ld)
    # # 傳統 端點加上簡算法的延伸長度
    # beam_ld_added = add_simple_ld(beam_ld_added)
    # beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl', beam_ld_added)
    # clock.time()

    # # 傳統斷筋
    # clock.time('傳統斷筋')
    # beam_3p_bar = cut_conservative(beam_ld_added, beam_3p)
    # clock.time()

    # # 三點斷筋
    # clock.time('三點斷筋')
    # beam_3p = cut_optimization(beam_ld_added, beam_3p)
    # clock.time()

    # # 輸出成表格
    # clock.time('輸出成表格')
    # beam_3p.to_excel(writer, '三點斷筋')
    # beam_3p_bar.to_excel(writer, '傳統斷筋')
    # beam_ld_added.to_excel(writer, 'beam_ld_added')
    # writer.save()
    # clock.time()


if __name__ == "__main__":
    # first_run()
    # second_run()
    first_full_run()
