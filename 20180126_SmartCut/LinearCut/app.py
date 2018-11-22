import os
import sys

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.pkl import load_pkl
from utils.Clock import Clock
from output_table import init_beam_3p, init_beam_name
from stirrups import calc_sturrups
from bar_size_num import calc_db_by_a_beam, calc_db_by_frame
from bar_con import cut_conservative, add_simple_ld
from bar_opti import calc_ld, add_ld, cut_optimization


# 不管是物件導向設計還是函數式編程 只要能解決問題的就是好方法
# 現在還只是看的不爽 所以並沒有造成問題
# 物件導向是對於真實世界的物體的映射
# 函數式編程是對於資料更好的操控

clock = Clock()


def first_run():
    clock.time()

    writer = pd.ExcelWriter(SCRIPT_DIR + '/dataset/first_run.xlsx')
    beam_name = init_beam_name()
    beam_name.to_excel(writer, '梁名編號')
    writer.save()

    clock.time()


def second_run():
    # second run
    writer = pd.ExcelWriter(SCRIPT_DIR + '/second_run.xlsx')

    # 初始化輸出表格
    clock.time('初始化輸出表格')
    beam_3p = init_beam_3p()
    clock.time()

    # 計算箍筋
    clock.time('計算箍筋')
    beam_3p, beam_v = calc_sturrups(beam_3p)
    (beam_3p, beam_v) = load_pkl(SCRIPT_DIR + '/stirrups.pkl', (beam_3p, beam_v))
    clock.time()

    # 以一台梁為單位 計算主筋 second run
    clock.time()
    beam_v_m = calc_db_by_frame(beam_v)
    beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl', beam_v_m)
    clock.time('以一台梁為單位 計算主筋')

    # 計算延伸長度
    clock.time()
    beam_v_m_ld = calc_ld(beam_v_m)
    clock.time('計算延伸長度')

    # 加上延伸長度
    clock.time()
    beam_ld_added = add_ld(beam_v_m_ld)
    # 傳統 端點加上簡算法的延伸長度
    beam_ld_added = add_simple_ld(beam_ld_added)
    beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl', beam_ld_added)
    clock.time('加上延伸長度')

    # 傳統斷筋
    clock.time()
    beam_3p_bar = cut_conservative(beam_ld_added, beam_3p)
    clock.time('傳統斷筋')

    # 三點斷筋
    clock.time()
    beam_3p = cut_optimization(beam_ld_added, beam_3p)
    clock.time('三點斷筋')

    # 輸出成表格
    clock.time()
    beam_3p.to_excel(writer, '三點斷筋')
    beam_3p_bar.to_excel(writer, '傳統斷筋')
    beam_ld_added.to_excel(writer, 'beam_ld_added')
    writer.save()
    clock.time('輸出成表格')


if __name__ == "__main__":
    # first_run()
    second_run()
