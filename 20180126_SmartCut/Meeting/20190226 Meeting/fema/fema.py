"""
fema function
"""


def c1(T, T0):
    if T < 0.1:
        return 1.5
    if T >= T0:
        return 1.0
    return (1.5 - 1.0) / (T0 - 0.1) * (T - 0.1)
