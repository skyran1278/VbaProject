import os
import re
import time
import pickle

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from dataset_beam_design import load_beam_design
from dataset_e2k import load_e2k
from output_table import init_beam_3points_table
from stirrups import calc_sturrups
from pkl import load_pkl

dataset_dir = os.path.dirname(os.path.abspath(__file__))
save_file = dataset_dir + '/3pionts.xlsx'
stirrups_save_file = dataset_dir + '/stirrups.pkl'


# beam_3points_table = init_beam_3points_table()
# beam_3points_table = calc_sturrups(beam_3points_table)
# beam_3points_table = load_pkl(stirrups_save_file, beam_3points_table)
beam_3points_table = load_pkl(stirrups_save_file)

beam_3points_table.to_excel(save_file)