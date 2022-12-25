@echo off
setlocal ENABLEDELAYEDEXPANSION

set str=%~1
set strCopy=%str%

set "str=%str:ö=^&ouml;%"
set "str=%str:Ö=^&Ouml;%"
set "str=%str:ä=^&auml;%"
set "str=%str:Ä=^&Auml;%"
set "str=%str:ü=^&uuml;%"
set "str=%str:Ü=^&Uuml;%"
set "str=%str:õ=^&otilde;%"
set "str=%str:Õ=^&Otilde;%"
set "str=%str:š=^&scaron;%"
set "str=%str:Š=^&Scaron;%"
set "str=%str:ž=^&zcaron;%"
set "str=%str:Ž=^&Zcaron;%"

echo.
echo %str%
echo.

set position=0
set oumlAmount=0
set aumlAmount=0
set uumlAmount=0
set otildeAmount=0
set scaronAmount=0
set zcaronAmount=0

:count
    if "!strCopy:~%position%,1!"=="ö" set /a oumlAmount=oumlAmount+1
    if "!strCopy:~%position%,1!"=="Ö" set /a oumlAmount=oumlAmount+1
    if "!strCopy:~%position%,1!"=="ä" set /a aumlAmount=aumlAmount+1
    if "!strCopy:~%position%,1!"=="Ä" set /a aumlAmount=aumlAmount+1
    if "!strCopy:~%position%,1!"=="ü" set /a uumlAmount=uumlAmount+1
    if "!strCopy:~%position%,1!"=="Ü" set /a uumlAmount=uumlAmount+1
    if "!strCopy:~%position%,1!"=="õ" set /a otildeAmount=otildeAmount+1
    if "!strCopy:~%position%,1!"=="Õ" set /a otildeAmount=otildeAmount+1
    if "!strCopy:~%position%,1!"=="š" set /a scaronAmount=scaronAmount+1
    if "!strCopy:~%position%,1!"=="Š" set /a scaronAmount=scaronAmount+1
    if "!strCopy:~%position%,1!"=="ž" set /a zcaronAmount=zcaronAmount+1
    if "!strCopy:~%position%,1!"=="Ž" set /a zcaronAmount=zcaronAmount+1
    set /a position=position+1
    if "!strCopy:~%position%,1!" NEQ "" goto count

set /a total=%oumlAmount%+%aumlAmount%+%uumlAmount%+%otildeAmount%+%scaronAmount%+%zcaronAmount%

if %total% GTR 0 (echo Vahetatud:
echo.
if %oumlAmount% GTR 0 echo ö is %oumlAmount%
if %aumlAmount% GTR 0 echo ä is %aumlAmount%
if %uumlAmount% GTR 0 echo ü is %uumlAmount%
if %otildeAmount% GTR 0 echo õ is %otildeAmount%
if %scaronAmount% GTR 0 echo š is %scaronAmount%
if %zcaronAmount% GTR 0 echo ž is %zcaronAmount%
echo Kokku: %total%) else (echo Ei leidnud ühtegi täpitähte.)