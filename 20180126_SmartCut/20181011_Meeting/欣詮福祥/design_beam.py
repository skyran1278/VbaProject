import numpy as np
import matplotlib.pyplot as plt

from dataset_design import load_beam_design

green = np.array([26, 188, 156]) / 256
blue = np.array([52, 152, 219]) / 256
red = np.array([233, 88, 73]) / 256
orange = np.array([230, 126, 34]) / 256
gray = np.array([0.5, 0.5, 0.5])
background = np.array([247, 247, 247]) / 256


class Beam():
    def __init__(self):
        pass

    def funcname(self, parameter_list):
        pass


beam_design = load_beam_design()

beam_story = np.unique(beam_design['Story'])

np.savetxt('')
