import os

import pandas as pd
import numpy as np

from utils.pkl import load_pkl
from utils.Clock import Clock

from output_table import init_beam_3p, init_beam_name
from stirrups import calc_sturrups
from bar_1st import calc_db_by_a_beam
from bar_con import cut_conservative
from bar_opti import calc_ld, add_ld, cut_optimization

dataset_dir = os.path.dirname(os.path.abspath(__file__))
# save_file = dataset_dir + '/3pionts.xlsx'

clock = Clock()

writer = pd.ExcelWriter(dataset_dir + '/dataset/first_run.xlsx')

# 不管是物件導向設計還是函數式編程 只要能解決問題的就是好方法
# 現在還只是看的不爽 所以並沒有造成問題
# 物件導向是對於真實世界的物體的映射
# 函數式編程是對於資料更好的操控

# 初始化輸出表格
beam_3p = init_beam_3p()
beam_name = init_beam_name()

# 計算箍筋
beam_3p, beam_v = calc_sturrups(beam_3p)
(beam_3p, beam_v) = load_pkl(dataset_dir + '/stirrups.pkl', (beam_3p, beam_v))

# 以一根梁為單位 計算主筋
beam_v_m = calc_db_by_a_beam(beam_v)
beam_v_m = load_pkl(dataset_dir + '/beam_v_m.pkl', beam_v_m)

# 計算延伸長度
beam_v_m_ld = calc_ld(beam_v_m)

# 加上延伸長度
beam_ld_added = add_ld(beam_v_m_ld)
beam_ld_added = load_pkl(dataset_dir + '/beam_ld_added.pkl', beam_ld_added)

# 傳統斷筋
beam_3p_bar = cut_conservative(beam_v_m, beam_3p)

# 三點斷筋
beam_3p = cut_optimization(beam_ld_added, beam_3p)

# 輸出成表格
beam_3p.to_excel(writer, '三點斷筋')
beam_3p_bar.to_excel(writer, '傳統斷筋')
beam_name.to_excel(writer, '梁名編號')
writer.save()
