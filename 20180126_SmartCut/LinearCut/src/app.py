""" app control """
import time

import pandas as pd

from const import OUTPUT_DIR, ETABS_DESIGN_PATH, E2K_PATH, BEAM_NAME_PATH
from utils.pkl import load_pkl
from utils.execution_time import Execution

from data.dataset_etabs_design import load_beam_design
from data.dataset_e2k import load_e2k
from data.dataset_beam_name import load_beam_name

from components.init_beam import init_beam, init_beam_name, add_and_alter_beam_id
from components.stirrups import calc_stirrups
from components.bar_size_num import calc_db
from components.bar_ld import calc_ld, add_ld
from components.bar_traditional import cut_traditional
from components.bar_cut import cut_optimization

# 不管是物件導向設計還是函數式編程 只要能解決問題的就是好方法
# 現在還只是看的不爽 所以並沒有造成問題
# 物件導向是對於真實世界的物體的映射
# 函數式編程是對於資料更好的操控


def cut_by_beam(moment=3, shear=False):
    """ run by beam, no need beam name ID"""
    execution = Execution()

    # get input data
    e2k = load_e2k(E2K_PATH, E2K_PATH + '.pkl')
    etabs_design = load_beam_design(
        ETABS_DESIGN_PATH, ETABS_DESIGN_PATH + '.pkl')
    beam_name = load_beam_name(BEAM_NAME_PATH, BEAM_NAME_PATH + '.pkl')

    # output path
    writer = pd.ExcelWriter(
        OUTPUT_DIR + '/' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ' SmartCut.xlsx')

    # 初始化輸出表格
    execution.time('初始化輸出表格')
    beam_name_empty = init_beam_name(etabs_design)
    beam_traditional = init_beam(etabs_design, e2k)
    beam = init_beam(etabs_design, e2k, moment=moment, shear=shear)
    execution.time()

    # 計算箍筋
    execution.time('計算箍筋')
    beam, dh_design = calc_stirrups(beam, etabs_design)
    beam_traditional, _ = calc_stirrups(beam_traditional, etabs_design)
    (beam, beam_traditional, dh_design) = load_pkl(
        OUTPUT_DIR + '/dh_design.pkl', (beam, beam_traditional, dh_design))
    execution.time()

    # 以一根梁為單位 計算主筋
    execution.time('以一根梁為單位 計算主筋')
    db_design = calc_db('BayID', dh_design, e2k)
    db_design = load_pkl(OUTPUT_DIR + '/db_design.pkl', db_design)
    execution.time()

    # 計算延伸長度
    execution.time('計算延伸長度')
    ld_design = calc_ld(db_design, e2k)
    execution.time()

    # 加上延伸長度
    execution.time('加上延伸長度')
    ld_design = add_ld(ld_design, 'Ld')
    ld_design = load_pkl(OUTPUT_DIR + '/ld_design.pkl', ld_design)
    execution.time()

    # 傳統斷筋
    execution.time('傳統斷筋')
    beam_traditional = cut_traditional(beam_traditional, ld_design)
    execution.time()

    # 多點斷筋
    execution.time('多點斷筋')
    beam = cut_optimization(moment, beam, ld_design)
    execution.time()

    # 輸出成表格
    execution.time('輸出成表格')
    beam_name_empty.to_excel(writer, '梁名編號')
    beam.to_excel(writer, '多點斷筋')
    beam_traditional.to_excel(writer, '傳統斷筋')
    ld_design.to_excel(writer, 'etabs_design')
    writer.save()
    execution.time()


def cut_by_frame(moment=3, shear=False):
    """ run by frame, need beam name ID"""
    execution = Execution()

    # get input data
    e2k = load_e2k(E2K_PATH, E2K_PATH + '.pkl')
    etabs_design = load_beam_design(
        ETABS_DESIGN_PATH, ETABS_DESIGN_PATH + '.pkl')
    beam_name = load_beam_name(BEAM_NAME_PATH, BEAM_NAME_PATH + '.pkl')

    # output path
    writer = pd.ExcelWriter(
        OUTPUT_DIR + '/' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + ' SmartCut.xlsx')

    # 初始化輸出表格
    execution.time('初始化輸出表格')
    beam_traditional = init_beam(etabs_design, e2k)
    beam = init_beam(etabs_design, e2k, moment=moment, shear=shear)
    # no change tradition beam id
    beam, etabs_design = add_and_alter_beam_id(
        beam, beam_name, etabs_design)
    execution.time()

    # 計算箍筋
    execution.time('計算箍筋')
    beam, dh_design = calc_stirrups(beam, etabs_design)
    beam_traditional, _ = calc_stirrups(beam_traditional, etabs_design)
    (beam, beam_traditional, dh_design) = load_pkl(
        OUTPUT_DIR + '/dh_design.pkl', (beam, beam_traditional, dh_design))
    execution.time()

    # 以一台梁為單位 計算主筋
    execution.time('以一台梁為單位 計算主筋')
    db_design = calc_db('FrameID', dh_design, e2k)
    db_design = load_pkl(OUTPUT_DIR + '/db_design.pkl', db_design)
    execution.time()

    # 計算延伸長度
    execution.time('計算延伸長度')
    ld_design = calc_ld(db_design, e2k)
    execution.time()

    # 加上延伸長度
    execution.time('加上延伸長度')
    ld_design = add_ld(ld_design, 'Ld')
    ld_design = load_pkl(OUTPUT_DIR + '/ld_design.pkl', ld_design)
    execution.time()

    # 傳統斷筋
    execution.time('傳統斷筋')
    beam_traditional = cut_traditional(beam_traditional, ld_design)
    execution.time()

    # 多點斷筋
    execution.time('多點斷筋')
    beam = cut_optimization(moment, beam, ld_design)
    execution.time()

    # 輸出成表格
    execution.time('輸出成表格')
    beam.to_excel(writer, '多點斷筋')
    beam_traditional.to_excel(writer, '傳統斷筋')
    ld_design.to_excel(writer, 'etabs_design')
    beam_name.to_excel(writer, '梁名編號')
    writer.save()
    execution.time()


if __name__ == "__main__":
    cut_by_frame()
    # cut_by_beam()
