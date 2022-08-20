@echo off


echo STAVES
echo..................
:ini

echo 0 - Todos
echo 1 - Anel 04
echo 2 - Anel 06
echo 3 - Anel 08
echo 4 - Anel 09
echo 5 - Anel 10
echo 6 - Anel 11
echo 7 - Anel 13
echo.
set /p dispositivo="Qual peca[0, 1, 2, 3, 4, 5, 6, 7]? "


if %dispositivo% == 0 (
	goto all
) else if %dispositivo% == 1 ( 
	goto Anel04
) else if %dispositivo% == 2 ( 
	goto Anel06
) else if %dispositivo% == 3 ( 
	goto Anel08
) else if %dispositivo% == 4 ( 
	goto Anel09
) else if %dispositivo% == 5 ( 
	goto Anel10
) else if %dispositivo% == 6 ( 
	goto Anel11
) else if %dispositivo% == 7 ( 
	goto Anel13
) else (
	echo opcao invalida... Tente novamente
	pause >nul
	cls
	goto :ini
)

:All
:Anel04

	set str="st01", "st02", "st03", "st04", "st05", "st06", "st07", "st08", "st09", "st10", "st11", "st12", "st13", "st14", "st15", "st16", "st17", "st18", "st19", "st20", "st21", "st22"
	set /a n=22
	pushd "IR\Anel04"
	call :processa
	popd
	pushd "Tratadas\Anel04"
	call :processa
	popd
	if %dispositivo% neq 0 (
		goto FIM
	)
	
:Anel06
	set str="st01", "st03", "st05", "st07", "st09", "st11", "st13", "st15", "st17", "st19", "st21", "st23", "st25", "st27", "st29", "st31"
	set /a n=16
	pushd "IR\Anel06"
	call :processa
	popd
	pushd "Tratadas\Anel06"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)

:Anel08
	set str="st01", "st03", "st05", "st07", "st09", "st11", "st13", "st15", "st17", "st19", "st21", "st23", "st25", "st27", "st29", "st31"
	set /a n=16
	pushd "IR\Anel08"
	call :processa
	popd
	pushd "Tratadas\Anel08"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)
	
:Anel09
	set str="st01", "st03", "st05", "st07", "st09", "st11", "st13", "st15", "st17", "st19", "st21", "st23", "st25", "st27", "st29", "st31"
	set /a n=16
	pushd "IR\Anel09"
	call :processa
	popd
	pushd "Tratadas\Anel09"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)

:Anel10
	set str="st02", "st04", "st06", "st08", "st10", "st12", "st14", "st16", "st18", "st20", "st22", "st24", "st26", "st28"
	set /a n=14
	pushd "IR\Anel10"
	call :processa
	popd
	pushd "Tratadas\Anel10"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)

:Anel11
	set str="st01", "st03", "st05", "st07", "st09", "st11", "st13", "st15", "st17", "st19", "st21", "st23", "st25", "st27"
	set /a n=14
	pushd "IR\Anel11"
	call :processa
	popd
	pushd "Tratadas\Anel11"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)
	
:Anel13
	set str="st01", "st04", "st08", "st11", "st15", "st18", "st22", "st25"
	set /a n=8
	pushd "IR\Anel13"
	call :processa
	popd
	pushd "Tratadas\Anel13"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)

:FIM
	echo.
	echo Renomeacao completa, favor conferir!!!!
	pause >nul
	goto :EOF


:processa
rem armazena o nomes na str
set /a cont=0
setlocal ENABLEDELAYEDEXPANSION
for %%s in (%str%) do (
	set nomecerto[!cont!]=%%~s
	set /a cont+=1
)

rem armazena o nome das imagens na pasta
set /a cont2=0
for %%j in (*.jpg) do (
	set nomeantigo[!cont2!]=%%j
	set /a cont2+=1
) 

rem renomea um pro outro

set /a cont3=0
:loop
	ren !nomeantigo[%cont3%]! !nomecerto[%cont3%]!.JPG
::	echo !nomeantigo[%cont3%]! !nomecerto[%cont3%]!.JPG
	set /a cont3+=1

if %cont3% lss %n% ( goto loop )
 
endlocal