import os
import sys

import pandas as pd
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(SCRIPT_DIR, os.path.pardir))

from utils.Clock import Clock
from utils.pkl import load_pkl

from output_table import init_beam_3p, init_beam_name
from stirrups import calc_sturrups
from bar_2nd import calc_db_by_frame
from bar_con import cut_conservative
from bar_opti import calc_ld, add_ld, cut_optimization


clock = Clock()
clock.time()
writer = pd.ExcelWriter(SCRIPT_DIR + '/second_run.xlsx')

# 初始化輸出表格
beam_3p = init_beam_3p()
beam_name = init_beam_name()

# 計算箍筋
beam_3p, beam_v = calc_sturrups(beam_3p)
(beam_3p, beam_v) = load_pkl(SCRIPT_DIR + '/stirrups.pkl', (beam_3p, beam_v))

# 以一個 frame 為單位 計算主筋
beam_v_m = calc_db_by_frame(beam_v)
beam_v_m = load_pkl(SCRIPT_DIR + '/beam_v_m.pkl', beam_v_m)

# 計算延伸長度
beam_v_m_ld = calc_ld(beam_v_m)

# 加上延伸長度
beam_ld_added = add_ld(beam_v_m_ld)
beam_ld_added = load_pkl(SCRIPT_DIR + '/beam_ld_added.pkl', beam_ld_added)

# 傳統斷筋
beam_3p_con = cut_conservative(beam_v_m, beam_3p)

# 三點斷筋
beam_3p = cut_optimization(beam_ld_added, beam_3p)

# 輸出成表格
beam_ld_added.to_excel(writer, 'beam_ld_added')
beam_3p.to_excel(writer, '三點斷筋')
beam_3p_con.to_excel(writer, '傳統斷筋')
# beam_name.to_excel(writer, '梁名編號')
writer.save()
clock.time()
