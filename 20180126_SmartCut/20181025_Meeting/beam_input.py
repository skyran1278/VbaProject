import os
import numpy as np
import matplotlib.pyplot as plt

from dataset_design import load_beam_design

green = np.array([26, 188, 156]) / 256
blue = np.array([52, 152, 219]) / 256
red = np.array([233, 88, 73]) / 256
orange = np.array([230, 126, 34]) / 256
gray = np.array([0.5, 0.5, 0.5])
background = np.array([247, 247, 247]) / 256

dataset_dir = os.path.dirname(os.path.abspath(__file__))


class Beam():
    def __init__(self):
        pass

    def funcname(self, parameter_list):
        pass


beam_design = load_beam_design()

beam_story, beam_index = np.unique(beam_design['Story'], return_index=True)

# arr1inds = beam_index.argsort()

# np.sort(beam_story, order=beam_index)

np.savetxt(dataset_dir + '/input.txt', beam_story[beam_index.argsort()],
           fmt='%s', header='Story, TOP Db, BOT Db')

print('Done!')
