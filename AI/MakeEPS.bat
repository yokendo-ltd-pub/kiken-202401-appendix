@echo off
setlocal enabledelayedexpansion
cd %~dp0
call :ShowProgress
del *.pdf *.eps *.svg >nul 2>&1
call :ShowProgress
lualatex -halt-on-error eq >nul 2>&1
if errorlevel 1 (
  echo .
  echo TeX入力データに間違いがあるようです。
  start %~dp0
  pause
  exit /b 1
)

call :ShowProgress
dvipdfmx eq.dvi >nul 2>&1
call :ShowProgress
rungs -dNOCACHE -dNOPAUSE -dBATCH -dNoOutputFonts -sDEVICE=eps2write -sOutputFile=eq.eps eq.pdf >nul 2>&1
call :ShowProgress
rungs -dNOCACHE -dNOPAUSE -dBATCH -dEPSCrop -sDEVICE=pdfwrite -sOutputFile=eq.pdf eq.eps >nul 2>&1
call :ShowProgress
dvisvgm -E eq.eps >nul 2>&1
call :ShowProgress
del *.dvi *.log *.aux

exit /b 0

:ShowProgress
if not defined count (
  set /p Bar=ファイル変換中.< nul
) else (
  set /p Bar=.< nul
)
set /a count+=1
exit /b