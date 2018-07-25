import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

greenColor = np.array([26, 188, 156]) / 256
blueColor = np.array([52, 152, 219]) / 256
redColor = np.array([233, 88, 73]) / 256
grayColor = [0.5, 0.5, 0.5]

plt.figure()
plt.plot([0, 1/3, 1/3, 2/3, 2/3, 1], [100, 100,
                                      80, 80, 100, 100], c=grayColor)
plt.plot([0, 0.4, 0.4, 0.6, 0.6, 1], [
         100, 100, 50, 50, 100, 100], c=greenColor, ls='-.')

plt.figure()
plt.plot([0, 1/3, 1/3, 2/3, 2/3, 1], [100, 100,
                                      80, 80, 100, 100], c=grayColor)
plt.plot([0, 0.2, 0.2, 0.8, 0.8, 1], [
         100, 100, 80, 80, 100, 100], c=greenColor, ls='-.')
plt.show()
# sns.lineplot([0, 1/3, 1/3, 2/3, 2/3, 1], [100, 100, 80, 80, 100, 100])
