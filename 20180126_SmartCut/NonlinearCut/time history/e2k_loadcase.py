"""
generate function and loadcase e2k with peernga data
"""
import shlex


def put_timehistorys(time_historys, peernga_folder):
    """
    put NPTS, DT
    """
    for time_history in time_historys:
        with open(f'{peernga_folder}//{time_history}.AT2', encoding='big5') as f:
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
            f'FUNCTION "{time_history}"  HEADERLINES 4  POINTSPERLINE 5  FORMAT "FREE"\n'
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


def post_pushover_loadcases(displacement, mode_shapes, initial_condition):
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
                f'MONITOREDDISPL  "Joint"  DISPLMAG  {displacement} DOF  "U1"  JOINT  "1"  "RF"\n'
            ),
            f'LOADCASE "{name}"  RESULTSSAVED  "Multiple"  MINSAVED  10 MAXSAVED  200\n',
            f'LOADCASE "{name}"  USEEVENTSTEPPING  "Yes"  MAXITERCS  4 MAXITERNR  10\n',
        )

    loadcases = []

    loadcases.extend(
        load('pushover x', f'LOADCASE "pushover x"  LOADPAT  "EQX"  SF  1\n'))

    loadcases.extend(load(
        'MMC',
        *[f'LOADCASE "MMC"  MODE  {mode} SF  {factor}\n' for mode, factor in enumerate(mode_shapes, 1)]
    ))

    for mode, factor in enumerate(mode_shapes, 1):
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

    peernga_folder = script_folder + '\\PEERNGARecords_Unscaled'

    direction = 'U1'

    initial_condition = '1.0DL + 0.5LL'
    dead_load = 'DL'
    live_load = 'LL'

    mode_shapes = [0.8, 0.1, 0.1]

    displacement = 1

    # different by model
    period = [0.039, 0.039 / 10]

    fectors = [1, 2, 3]

    time_historys = {
        'RSN169_IMPVALL.H_H-DLT262': {
            'FACTORS': fectors
        },
        'RSN953_NORTHR_MUL009':  {
            'FACTORS': fectors
        },
    }

    put_timehistorys(time_historys, peernga_folder)

    with open(script_folder + '/e2k_loadcase.e2k', mode='w', encoding='big5') as f:
        f.writelines(post_functions(time_historys, peernga_folder))
        f.write('\n\n\n')

        f.writelines(post_previous_loadcases(
            initial_condition, dead_load, live_load))
        f.writelines(post_pushover_loadcases(
            displacement, mode_shapes, initial_condition))

        f.writelines(post_timehistorys_loadcases(
            time_historys, period, initial_condition, direction))


if __name__ == "__main__":
    main()
