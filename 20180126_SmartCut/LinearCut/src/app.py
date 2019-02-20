""" app control """
import time

import pandas as pd

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


def cut_by_beam(const, moment=3, shear=False):
    """ run by beam, no need beam name ID"""
    e2k_path, etabs_design_path, output_dir = const[
        'e2k_path'], const['etabs_design_path'], const['output_dir']

    execution = Execution()

    # get input data
    e2k = load_e2k(e2k_path, e2k_path + '.pkl')
    etabs_design = load_beam_design(
        etabs_design_path, etabs_design_path + '.pkl')

    # output path
    writer = pd.ExcelWriter(
        output_dir + '/' + time.strftime("%Y%m%d %H%M%S", time.localtime()) + ' SmartCut.xlsx')

    # 初始化輸出表格
    execution.time('Initialize Output Table')
    beam_name_empty = init_beam_name(etabs_design)
    beam_traditional = init_beam(etabs_design, e2k)
    beam = init_beam(etabs_design, e2k, moment=moment, shear=shear)
    execution.time()

    # 計算箍筋
    execution.time('Calculate Stirrup Size and Spacing')
    beam, dh_design = calc_stirrups(beam, etabs_design, const)
    beam_traditional, _ = calc_stirrups(beam_traditional, etabs_design, const)
    (beam, beam_traditional, dh_design) = load_pkl(
        output_dir + '/dh_design.pkl', (beam, beam_traditional, dh_design))
    execution.time()

    # 以一根梁為單位 計算主筋
    execution.time('Calculate Rebar Size and Number by Beam')
    db_design = calc_db('BayID', dh_design, e2k, const)
    db_design = load_pkl(output_dir + '/db_design.pkl', db_design)
    execution.time()

    # 計算延伸長度
    execution.time('Calculate Ld')
    ld_design = calc_ld(db_design, e2k, const)
    execution.time()

    # 加上延伸長度
    execution.time('Add Ld')
    ld_design = add_ld(ld_design, 'Ld', const['rebar'])
    ld_design = load_pkl(output_dir + '/ld_design.pkl', ld_design)
    execution.time()

    # 傳統斷筋
    execution.time('Traditional Cut')
    beam_traditional = cut_traditional(
        beam_traditional, ld_design, const['rebar'])
    execution.time()

    # 多點斷筋
    execution.time('Multi Smart Cut')
    beam = cut_optimization(moment, beam, ld_design, const)
    execution.time()

    # 輸出成表格
    execution.time('Output Result')
    beam_name_empty.to_excel(writer, '梁名編號')
    beam.to_excel(writer, '多點斷筋')
    beam_traditional.to_excel(writer, '傳統斷筋')
    ld_design.to_excel(writer, 'etabs_design')
    writer.save()
    execution.time()


def cut_by_frame(const, moment=3, shear=False):
    """ run by frame, need beam name ID"""
    e2k_path, etabs_design_path, beam_name_path, output_dir = const[
        'e2k_path'], const['etabs_design_path'], const['beam_name_path'], const['output_dir']

    execution = Execution()

    # get input data
    e2k = load_e2k(e2k_path, e2k_path + '.pkl')
    etabs_design = load_beam_design(
        etabs_design_path, etabs_design_path + '.pkl')
    beam_name = load_beam_name(beam_name_path, beam_name_path + '.pkl')

    # output path
    writer = pd.ExcelWriter(
        output_dir + '/' + time.strftime("%Y%m%d %H%M%S", time.localtime()) + ' SmartCut.xlsx')

    # 初始化輸出表格
    execution.time('Initialize Output Table')
    beam_traditional = init_beam(etabs_design, e2k)
    beam = init_beam(etabs_design, e2k, moment=moment, shear=shear)
    # no change tradition beam id
    beam, etabs_design = add_and_alter_beam_id(
        beam, beam_name, etabs_design)
    execution.time()

    # 計算箍筋
    execution.time('Calculate Stirrup Size and Spacing')
    beam, dh_design = calc_stirrups(beam, etabs_design, const)
    beam_traditional, _ = calc_stirrups(beam_traditional, etabs_design, const)
    (beam, beam_traditional, dh_design) = load_pkl(
        output_dir + '/dh_design.pkl', (beam, beam_traditional, dh_design))
    execution.time()

    # 以一台梁為單位 計算主筋
    execution.time('Calculate Rebar Size and Number by Frame')
    db_design = calc_db('FrameID', dh_design, e2k, const)
    db_design = load_pkl(output_dir + '/db_design.pkl', db_design)
    execution.time()

    # 計算延伸長度
    execution.time('Calculate Ld')
    ld_design = calc_ld(db_design, e2k, const)
    execution.time()

    # 加上延伸長度
    execution.time('Add Ld')
    ld_design = add_ld(ld_design, 'Ld', const['rebar'])
    ld_design = load_pkl(output_dir + '/ld_design.pkl', ld_design)
    execution.time()

    # 傳統斷筋
    execution.time('Traditional Cut')
    beam_traditional = cut_traditional(
        beam_traditional, ld_design, const['rebar'])
    execution.time()

    # 多點斷筋
    execution.time('Multi Smart Cut')
    beam = cut_optimization(moment, beam, ld_design, const)
    execution.time()

    # 輸出成表格
    execution.time('Output Result')
    beam.to_excel(writer, '多點斷筋')
    beam_traditional.to_excel(writer, '傳統斷筋')
    ld_design.to_excel(writer, 'etabs_design')
    beam_name.to_excel(writer, '梁名編號')
    writer.save()
    execution.time()


if __name__ == "__main__":
    from const import const as constants

    cut_by_frame(constants)
    cut_by_beam(constants)
