"""
generate function and loadcase e2k with peernga data
"""
import shlex

CONFIG = {
    # different by model
    'factors': [0.2, 0.5, 0.772, 1.029, 1.5, 2, 2.5],
    'modal_participating_mass': [0.8528, 0.106, 0.0335, 0.0077],
    'period': [0.094, 0.094 / 10],
    'displacement': 0.8,
}


def put_timehistorys(time_historys, peernga_folder):
    """
    put NPTS, DT
    """
    for time_history in time_historys:
        with open(f'{peernga_folder}/{time_history}.AT2') as f:
            words = shlex.split(f.readlines()[3])
            time_historys[time_history]['NPTS'] = int(words[1][:-1])
            time_historys[time_history]['DT'] = float(words[3])


def post_functions(time_historys, peernga_folder):
    """
    path folder same as time history

    @example
    FUNCTION "RSN169_IMPVALL.H_H-DLT262"  FUNCTYPE "HISTORY"
    FILE "path.AT2"  DATATYPE "EQUAL"  DT 0.01

    FUNCTION "RSN169_IMPVALL.H_H-DLT262"  HEADERLINES 4  POINTSPERLINE 5  FORMAT "FREE"
    """
    # post functions to list
    functions = []

    for time_history in time_historys:
        delta_t = time_historys[time_history]['DT']

        functions.append(
            f'FUNCTION "{time_history}"  FUNCTYPE "HISTORY"  '
            f'FILE "{peernga_folder}\\{time_history}.AT2"  '
            f'DATATYPE "EQUAL"  DT {delta_t}\n'
        )

        functions.append(
            f'FUNCTION "{time_history}"  HEADERLINES 4  POINTSPERLINE 1  FORMAT "FREE"\n'
        )

    return functions


def post_previous_loadcases(initial_condition, dead_load, live_load):
    """
    LOADCASE "1.0DL + 0.5LL"  TYPE  "Nonlinear Static"  INITCOND  "NONE"
    MODALCASE  "Modal"  MASSSOURCE  "Previous"

    LOADCASE "1.0DL + 0.5LL"  LOADPAT  "DL"  SF  1

    LOADCASE "1.0DL + 0.5LL"  LOADPAT  "LL"  SF  0.5

    LOADCASE "1.0DL + 0.5LL"  NLGEOMTYPE  "PDelta"

    LOADCASE "1.0DL + 0.5LL"  LOADCONTROL  "Full"  DISPLTYPE  "Monitored"
    MONITOREDDISPL  "Joint"  DISPLMAG  0 DOF  "U3"  JOINT  "1"  "RF"

    LOADCASE "1.0DL + 0.5LL"  RESULTSSAVED  "Multiple"  MINSAVED  1 MAXSAVED  200

    LOADCASE "1.0DL + 0.5LL"  USEEVENTSTEPPING  "Yes"  MAXITERCS  4 MAXITERNR  10

    """
    # pylint: disable=line-too-long
    return (
        f'LOADCASE "{initial_condition}"  TYPE  "Nonlinear Static"  INITCOND  "NONE"  MODALCASE  "Modal"  MASSSOURCE  "Previous"\n',
        f'LOADCASE "{initial_condition}"  LOADPAT  "{dead_load}"  SF  1\n',
        f'LOADCASE "{initial_condition}"  LOADPAT  "{live_load}"  SF  0.5\n',
        f'LOADCASE "{initial_condition}"  NLGEOMTYPE  "PDelta"\n',
        f'LOADCASE "{initial_condition}"  LOADCONTROL  "Full"  DISPLTYPE  "Monitored"  MONITOREDDISPL  "Joint"  DISPLMAG  0 DOF  "U3"  JOINT  "1"  "RF"\n'
        f'LOADCASE "{initial_condition}"  RESULTSSAVED  "Multiple"  MINSAVED  1 MAXSAVED  200\n'
        f'LOADCASE "{initial_condition}"  USEEVENTSTEPPING  "Yes"  MAXITERCS  4 MAXITERNR  10\n'
    )


def post_pushover_loadcases(displacement, modal_participating_mass, initial_condition, direction):
    """
    LOADCASE "pushover x"  TYPE  "Nonlinear Static"  INITCOND  "1.0DL + 0.5LL"
    MODALCASE  "Modal"  MASSSOURCE  "Previous"

    LOADCASE "pushover x"  LOADPAT  "EQX"  SF  1

    LOADCASE "pushover x"  NLGEOMTYPE  "PDelta"

    LOADCASE "pushover x"  LOADCONTROL  "Displacement"  DISPLTYPE  "Conjugate"
    MONITOREDDISPL  "Joint"  DISPLMAG  0.2 DOF  "U1"  JOINT  "1"  "RF"

    LOADCASE "pushover x"  RESULTSSAVED  "Multiple"  MINSAVED  10 MAXSAVED  200

    LOADCASE "pushover x"  USEEVENTSTEPPING  "Yes"  MAXITERCS  4 MAXITERNR  10

    LOADCASE "MMC"  TYPE  "Nonlinear Static"  INITCOND  "PUSHDLLL"
    MODALCASE  "Modal"  MASSSOURCE  "Previous"

    LOADCASE "MMC"  MODE  1 SF  0.835511

    LOADCASE "MMC"  MODE  2 SF  0.112514

    LOADCASE "MMC"  NLGEOMTYPE  "PDelta"

    LOADCASE "MMC"  LOADCONTROL  "Displacement"  DISPLTYPE  "Conjugate"
    MONITOREDDISPL  "Joint"  DISPLMAG  0.2 DOF  "U1"  JOINT  "1"  "RF"

    LOADCASE "MMC"  RESULTSSAVED  "Multiple"  MINSAVED  10 MAXSAVED  200

    LOADCASE "MMC"  USEEVENTSTEPPING  "Yes"  MAXITERCS  4 MAXITERNR  10

    LOADCASE "PUSHXP1"  TYPE  "Nonlinear Static"  INITCOND  "PUSHDLLL"
    MODALCASE  "Modal"  MASSSOURCE  "Previous"

    LOADCASE "PUSHXP1"  MODE  1 SF  1

    LOADCASE "PUSHXP1"  NLGEOMTYPE  "PDelta"

    LOADCASE "PUSHXP1"  LOADCONTROL  "Displacement"  DISPLTYPE  "Conjugate"
    MONITOREDDISPL  "Joint"  DISPLMAG  0.2 DOF  "U1"  JOINT  "1"  "RF"

    LOADCASE "PUSHXP1"  RESULTSSAVED  "Multiple"  MINSAVED  10 MAXSAVED  200

    LOADCASE "PUSHXP1"  USEEVENTSTEPPING  "Yes"  MAXITERCS  4 MAXITERNR  10

    """
    def load(name, *load_pattern):
        return (
            (
                f'LOADCASE "{name}"  TYPE  "Nonlinear Static"  '
                f'INITCOND  "{initial_condition}"  MODALCASE  "Modal"  MASSSOURCE  "Previous"\n'
            ),
            *load_pattern,
            f'LOADCASE "{name}"  NLGEOMTYPE  "PDelta"\n',
            (
                f'LOADCASE "{name}"  LOADCONTROL  "Displacement"  DISPLTYPE  "Conjugate"  '
                f'MONITOREDDISPL  "Joint"  DISPLMAG  {displacement} '
                f'DOF  "{direction}"  JOINT  "1"  "RF"\n'
            ),
            f'LOADCASE "{name}"  RESULTSSAVED  "Multiple"  MINSAVED  10 MAXSAVED  200\n',
            f'LOADCASE "{name}"  USEEVENTSTEPPING  "Yes"  MAXITERCS  4 MAXITERNR  10\n',
        )

    loadcases = []

    loadcases.extend(
        load('pushover x', f'LOADCASE "pushover x"  LOADPAT  "EQX"  SF  1\n'))

    loadcases.extend(load(
        'MMC',
        *[
            f'LOADCASE "MMC"  MODE  {mode} SF  {factor}\n'
            for mode, factor in enumerate(modal_participating_mass, 1)
        ]
    ))

    for mode, factor in enumerate(modal_participating_mass, 1):
        loadcases.extend(load(
            f'pushover x {mode}',
            f'LOADCASE "pushover x {mode}"  MODE  {mode} SF  {factor}\n'
        ))

    return loadcases


def post_timehistorys_loadcases(time_historys, period, initial_condition, direction):
    """
    loadcase

    @example
    LOADCASE "timehistory"  TYPE  "Nonlinear Direct Integration History"
    INITCOND  "1.0DL + 0.5LL"  MODALCASE  "Modal"  MASSSOURCE  "Previous"

    LOADCASE "timehistory"  ACCEL  "U1"  FUNC  "chichi_TCU052_max"  SF  1

    LOADCASE "timehistory"  NLGEOMTYPE  "PDelta"  NUMBEROUTPUTSTEPS  100 OUTPUTSTEPSIZE  0.1

    LOADCASE "timehistory"  PRODAMPTYPE  "Period"  T1  0.344 DAMP1  0.05 T2  0.088 DAMP2  0.05

    LOADCASE "timehistory"  MODALDAMPTYPE  "Constant"  CONSTDAMP  0.05
    CONSIDERMAXMODALFREQ  "Yes"  MAXCONSIDEREDMODALFREQ  100

    LOADCASE "timehistory"  USEEVENTSTEPPING  "No"
    """
    loadcases = []

    for time_history in time_historys:
        factors = time_historys[time_history]['FACTORS']
        number_output_steps = time_historys[time_history]['NPTS']
        delta_t = time_historys[time_history]['DT']

        for factor in factors:
            name = f'{time_history}-{factor}'

            # G to m/s2
            factor = factor * 9.81

            loadcases.append(
                f'LOADCASE "{name}"  TYPE  "Nonlinear Direct Integration History"  '
                f'INITCOND  "{initial_condition}"  MODALCASE  "Modal"  MASSSOURCE  "Previous"\n'
            )

            loadcases.append(
                f'LOADCASE "{name}"  ACCEL  "{direction}"  '
                f'FUNC  "{time_history}"  SF  {factor}\n'
            )
            loadcases.append(
                f'LOADCASE "{name}"  NLGEOMTYPE  "PDelta"  NUMBEROUTPUTSTEPS  {number_output_steps} '
                f'OUTPUTSTEPSIZE  {delta_t}\n'
            )
            loadcases.append(
                f'LOADCASE "{name}"  PRODAMPTYPE  "Period"  T1  {period[0]} DAMP1  0.05 '
                f'T2  {period[1]} DAMP2  0.05\n'
            )
            loadcases.append(
                f'LOADCASE "{name}"  MODALDAMPTYPE  "Constant"  CONSTDAMP  0.05 '
                f'CONSIDERMAXMODALFREQ  "Yes"  MAXCONSIDEREDMODALFREQ  100 \n'
            )
            loadcases.append(f'LOADCASE "{name}"  USEEVENTSTEPPING  "No"\n')

    return loadcases


def main():
    """
    test
    """
    import os

    # global
    script_folder = os.path.dirname(os.path.abspath(__file__))

    peernga_folder = script_folder + '/PEERNGARecords_Normalized'

    direction = 'U1'

    initial_condition = '1.0DL + 0.5LL'
    dead_load = 'DL'
    live_load = 'LL'

    # different by model
    displacement = CONFIG['displacement']

    modal_participating_mass = CONFIG['modal_participating_mass']

    period = CONFIG['period']

    factors = CONFIG['factors']

    time_historys = {
        # 'RSN125_FRIULI.A_A-TMZ000': {'FACTORS': factors},
        # 'RSN767_LOMAP_G03000': {'FACTORS': factors},
        # 'RSN1148_KOCAELI_ARE000': {'FACTORS': factors},
        # 'RSN1602_DUZCE_BOL000': {'FACTORS': factors},
        # 'RSN1111_KOBE_NIS090': {'FACTORS': factors},
        # 'RSN1633_MANJIL_ABBAR--L': {'FACTORS': factors},
        # 'RSN725_SUPER.B_B-POE270': {'FACTORS': factors},
        # 'RSN68_SFERN_PEL180': {'FACTORS': factors},
        # 'RSN960_NORTHR_LOS270': {'FACTORS': factors},
        # 'RSN1485_CHICHI_TCU045-N': {'FACTORS': factors},
        'RSN68_SFERN_PEL090':  {'FACTORS': factors},
        'RSN125_FRIULI.A_A-TMZ270':  {'FACTORS': factors},
        'RSN169_IMPVALL.H_H-DLT262':  {'FACTORS': factors},
        'RSN174_IMPVALL.H_H-E11230':  {'FACTORS': factors},
        'RSN721_SUPER.B_B-ICC090':  {'FACTORS': factors},
        'RSN725_SUPER.B_B-POE360':  {'FACTORS': factors},
        'RSN752_LOMAP_CAP000':  {'FACTORS': factors},
        'RSN767_LOMAP_G03090':  {'FACTORS': factors},
        'RSN848_LANDERS_CLW-TR':  {'FACTORS': factors},
        'RSN900_LANDERS_YER270':  {'FACTORS': factors},
        'RSN953_NORTHR_MUL279':  {'FACTORS': factors},
        'RSN960_NORTHR_LOS000':  {'FACTORS': factors},
        'RSN1111_KOBE_NIS000':  {'FACTORS': factors},
        'RSN1116_KOBE_SHI000':  {'FACTORS': factors},
        'RSN1148_KOCAELI_ARE090':  {'FACTORS': factors},
        'RSN1158_KOCAELI_DZC180':  {'FACTORS': factors},
        'RSN1244_CHICHI_CHY101-N':  {'FACTORS': factors},
        'RSN1485_CHICHI_TCU045-E':  {'FACTORS': factors},
        'RSN1602_DUZCE_BOL090':  {'FACTORS': factors},
        'RSN1633_MANJIL_ABBAR--T':  {'FACTORS': factors},
        'RSN1787_HECTOR_HEC090':  {'FACTORS': factors},
    }

    put_timehistorys(time_historys, peernga_folder)

    with open(script_folder + '/e2k_loadcase.e2k', mode='w', encoding='big5') as f:
        f.writelines(post_functions(time_historys, peernga_folder))
        f.write('\n\n\n')

        f.writelines(post_previous_loadcases(
            initial_condition, dead_load, live_load))
        f.writelines(post_pushover_loadcases(
            displacement, modal_participating_mass, initial_condition, direction))

        f.writelines(post_timehistorys_loadcases(
            time_historys, period, initial_condition, direction))


if __name__ == "__main__":
    main()
