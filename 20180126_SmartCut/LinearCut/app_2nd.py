import os
import re
import time
import pickle

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset.dataset_beam_design import load_beam_design
from dataset.dataset_e2k import load_e2k
from output_table import init_beam_3points_table
from stirrups import calc_sturrups
from utils.pkl import load_pkl

dataset_dir = os.path.dirname(os.path.abspath(__file__))
save_file = dataset_dir + '/3pionts.xlsx'
stirrups_save_file = dataset_dir + '/stirrups.pkl'

beam_3points_table = init_beam_3points_table()
beam_3points_table, beam_design_table_with_stirrups = calc_sturrups(
    beam_3points_table)
(beam_3points_table, beam_design_table_with_stirrups) = load_pkl(
    stirrups_save_file, (beam_3points_table, beam_design_table_with_stirrups))
# (beam_3points_table, beam_design_table_with_stirrups) = load_pkl(stirrups_save_file)

beam_design_table_with_stirrups.to_excel(save_file)
