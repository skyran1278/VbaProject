

#   FUNCTION "chichi_TCU052_max"  FUNCTYPE "HISTORY"  FILE "D:\GitHub\VbaProject\20180126_SmartCut\NonlinearCut\time history\chichi_TCU052_max.txt"  DATATYPE "T&V"
#   FUNCTION "chichi_TCU052_max"  POINTSPERLINE 1  FORMAT "FREE"

#   LOADCASE "timehistory"  TYPE  "Nonlinear Direct Integration History"  INITCOND  "PUSHDLLL"  MODALCASE  "Modal"  MASSSOURCE  "Previous"
#   LOADCASE "timehistory"  ACCEL  "U1"  FUNC  "chichi_TCU052_max"  SF  1
#   LOADCASE "timehistory"  NUMBEROUTPUTSTEPS  100 OUTPUTSTEPSIZE  0.1
#   LOADCASE "timehistory"  PRODAMPTYPE  "Period"  T1  0.344 DAMP1  0.05 T2  0.088 DAMP2  0.05
#   LOADCASE "timehistory"  MODALDAMPTYPE  "None"
#   LOADCASE "timehistory"  USEEVENTSTEPPING  "No"
