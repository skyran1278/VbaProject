"""
rebar constant and functions
because etabs defined rebars is not avaliable to Taiwan
"""

# unit: m, m2
REBARS = {
    ('#3', 'DB'): 0.00953,
    ('#4', 'DB'): 0.0127,
    ('#5', 'DB'): 0.0159,
    ('#6', 'DB'): 0.0191,
    ('#7', 'DB'): 0.0222,
    ('#8', 'DB'): 0.0254,
    ('#9', 'DB'): 0.0287,
    ('#10', 'DB'): 0.0322,
    ('#11', 'DB'): 0.0358,
    ('#12', 'DB'): 0.0394,
    ('#14', 'DB'): 0.0430,
    ('#16', 'DB'): 0.0502,
    ('#18', 'DB'): 0.0573,
    ('#3', 'AREA'): 0.00007133,
    ('#4', 'AREA'): 0.0001267,
    ('#5', 'AREA'): 0.0001986,
    ('#6', 'AREA'): 0.0002865,
    ('#7', 'AREA'): 0.0003871,
    ('#8', 'AREA'): 0.0005067,
    ('#9', 'AREA'): 0.0006469,
    ('#10', 'AREA'): 0.0008143,
    ('#11', 'AREA'): 0.001007,
    ('#12', 'AREA'): 0.001219,
    ('#14', 'AREA'): 0.001452,
    ('#16', 'AREA'): 0.001979,
    ('#18', 'AREA'): 0.002579,
}


def double_area(number):
    """
    double area, for stirrups\n
    number: string, rebar No. (ex: '#4')
    """
    return REBARS[(number, 'AREA')] * 2


def rebar_db(number):
    """ rebar db
    """
    return REBARS[(number, 'DB')] * 2


def rebar_area(number):
    """ rebar area
    """
    return REBARS[(number, 'AREA')] * 2
