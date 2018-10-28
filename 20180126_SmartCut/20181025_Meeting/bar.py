import os

from dataset.dataset_beam_design import load_beam_design
from dataset.dataset_e2k import load_e2k
from dataset.const import TOP_BAR, BOT_BAR, DB_SPACING
from stirrups import calc_sturrups
from output_table import init_beam_3points_table
from utils.pkl import load_pkl, create_pkl

dataset_dir = os.path.dirname(os.path.abspath(__file__))
save_file = dataset_dir + '/3pionts.xlsx'
stirrups_save_file = dataset_dir + '/stirrups.pkl'

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()
# beam_design_table = load_beam_design()
# beam_3points_table = init_beam_3points_table()
# beam_3points_table, beam_design_table_with_stirrups = calc_sturrups(beam_3points_table)
(beam_3points_table, beam_design_table) = load_pkl(stirrups_save_file)

dataset_const = {
    'TOP': TOP_BAR,
    'BOT': BOT_BAR
}

for BAR in ('TOP', 'BOT'):
    Bar = BAR.capitalize()
    i = 0
    beam_design_table = beam_design_table.assign(
        bar_size=dataset_const[BAR][i], bar_num=lambda x: rebars[dataset_const[BAR][i], 'AREA'] / x['As' + Bar])
    print(beam_design_table.head())
    for (Story, BayID), group in beam_design_table.groupby(['Story', 'BayID'], sort=False):
        pass
