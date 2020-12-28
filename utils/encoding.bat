@echo off

@REM https://ss64.com/nt/for_r.html
@REM https://superuser.com/questions/27060/batch-convert-encoding-in-files
for /r %~dp0 %%i in (*.vb) do (
  iconv -f cp950 -t utf-8 "%%i" > "%%i8765432"
  mv "%%i8765432" "%%i"
)
