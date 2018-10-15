import os
import re
import numpy as np

from dataset_beam_design import load_beam_design
from dataset_e2k import load_e2k

rebars, stories, point_coordinates, lines, materials, sections = load_e2k()

beam_design = load_beam_design()

print('Done')
