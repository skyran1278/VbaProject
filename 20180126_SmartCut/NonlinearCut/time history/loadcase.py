import os

# global
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

initial_condition = 'PUSHDLLL'

direction = 'U1'

# different by model
period1 = 0.344
period2 = 0.088

time_historys = {
    'RSN169_IMPVALL.H_H-DLT262': [1, 2, 3],
    'RSN953_NORTHR_MUL009': [1, 2]
}

# different by function

for time_history in time_historys:
    with open(f'{SCRIPT_DIR}//{time_history}.AT2', encoding='big5') as f:
        pass

with open(SCRIPT_DIR + '/loadcase.e2k', mode='w', encoding='big5') as f:
    for name in time_historys:
        post_function(f, name)


def post_function(f, name):
    """
    path folder same as time history
    """
    f.write(
        f'FUNCTION "{name}"  FUNCTYPE "HISTORY"  '
        f'{SCRIPT_DIR}//{name}.AT2"  '
        f'DATATYPE "EQUAL"  DT 0.01\n'
    )
    f.write(
        f'FUNCTION "{name}"  HEADERLINES 4  POINTSPERLINE 5  FORMAT "FREE"'
    )


def post_loadcase(f):
    name = f'{function}-{factor}'

    f.write(
        f'LOADCASE "{name}"  TYPE  "Nonlinear Direct Integration History"  '
        f'INITCOND  "{initial_condition}"  MODALCASE  "Modal"  MASSSOURCE  "Previous"\n'
    )

    f.write(
        f'LOADCASE "{name}"  ACCEL  "{direction}"  '
        f'FUNC  "{function}"  SF  {factor}\n'
    )
    f.write(
        f'LOADCASE "{name}"  NUMBEROUTPUTSTEPS  {output_step} '
        f'OUTPUTSTEPSIZE  {delta_t}\n'
    )
    f.write(
        f'LOADCASE "{name}"  PRODAMPTYPE  "Period"  T1  {period1} DAMP1  0.05 '
        f'T2  {period2} DAMP2  0.05\n'
    )
    f.write(f'LOADCASE "{name}"  MODALDAMPTYPE  "None"\n')
    f.write(f'LOADCASE "{name}"  USEEVENTSTEPPING  "No"\n')


#   FUNCTION "RSN169_IMPVALL.H_H-DLT262"  FUNCTYPE "HISTORY"  FILE "C:\Users\skyran\Downloads\PEERNGARecords_Unscaled\RSN169_IMPVALL.H_H-DLT262.AT2"  DATATYPE "EQUAL"  DT 0.01
#   FUNCTION "RSN169_IMPVALL.H_H-DLT262"  HEADERLINES 4  POINTSPERLINE 5  FORMAT "FREE"

#   FUNCTION "chichi_TCU052_max"  FUNCTYPE "HISTORY"  FILE "D:\GitHub\VbaProject\20180126_SmartCut\NonlinearCut\time history\chichi_TCU052_max.txt"  DATATYPE "T&V"
#   FUNCTION "chichi_TCU052_max"  POINTSPERLINE 1  FORMAT "FREE"

#   LOADCASE "timehistory"  TYPE  "Nonlinear Direct Integration History"  INITCOND  "PUSHDLLL"  MODALCASE  "Modal"  MASSSOURCE  "Previous"
#   LOADCASE "timehistory"  ACCEL  "U1"  FUNC  "chichi_TCU052_max"  SF  1
#   LOADCASE "timehistory"  NUMBEROUTPUTSTEPS  100 OUTPUTSTEPSIZE  0.1
#   LOADCASE "timehistory"  PRODAMPTYPE  "Period"  T1  0.344 DAMP1  0.05 T2  0.088 DAMP2  0.05
#   LOADCASE "timehistory"  MODALDAMPTYPE  "None"
#   LOADCASE "timehistory"  USEEVENTSTEPPING  "No"
