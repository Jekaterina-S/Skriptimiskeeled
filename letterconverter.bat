@echo off
setlocal enabledelayedexpansion

:: Check if a parameter is passed when invoking
if "%~1" equ "" (
    set /p "input=Sisestage lause: "
) else (
    set "input=%~1"
)

set "output="
set replaced=
set total=0
:: Calculating input lenght
call:strLen input inputlen

:: Going through all estonian umlaut letters
for %%U in (
    ä:auml; Ä:Auml;
    ö:ouml; Ö:Ouml;
    ü:uuml; Ü:Uuml;
    õ:otilde; Õ:Otilde;
    š:scaron; Š:Scaron;
    ž:zcaron; Ž:Zcaron;
) do (
    for /f "tokens=1-2 delims=:" %%A in ("%%U") do (
        set count=0
        :: Going through the input to find current umlaut, case sensitive
        for /l %%L in (0,1,!inputlen!) do (
            set char=!input:~%%L,1!
            if "!char!" equ "%%A" (
                set "char=&%%B;"
                set /a count+=1
                :: Adding maximum possible replacement character amount to lenght
                set /a inputlen+=8
            )
            set output=!output!!char!
        )
        :: Replacing letters and adding to total
        if !count! gtr 0 (
            set "input=!output!"
            set replaced=!replaced!%%A: !count! ;
            set /a total+=count
        )
        set "output="
    )
)

:: Output
echo.
echo !input!
echo.

if !total! equ 0 (
    echo Ei leidnud ühtegi täpitähte.
) else (
    echo Vahetatud:
    echo.
    for %%L in ("!replaced:;=" "!") do (
        if "%%~L" neq "" echo %%~L
    )
    echo.
    echo Kokku: %total%
)

exit /b

:: For calculating lenght of input string
:strLen
setlocal enabledelayedexpansion

:strLen_Loop
   if not "!%1:~%len%!"=="" set /A len+=1 & goto :strLen_Loop
(endlocal & set %2=%len%)
goto :eof
