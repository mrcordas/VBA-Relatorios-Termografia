@echo off


echo CONJUNTO PORTA VENTO
echo..................
echo Certifique-se das imagens ficar na ordem correta
echo.

:ini

echo 0 - Todos
echo 1 - Saida
echo 2 - DownLeg
echo 3 - Joelho
echo 4 - Nariz
echo.
set /p dispositivo="Qual peca[0, 1, 2, 3, 4]? "

set str="vt01", "vt02", "vt03", "vt04", "vt05", "vt06", "vt07", "vt08", "vt09", "vt10", "vt11", "vt12", "vt13", "vt14", "vt15", "vt16", "vt17", "vt18", "vt19", "vt20", "vt21", "vt22"
set /a n=44

if %dispositivo% == 0 (
	goto all
) else if %dispositivo% == 1 ( 
	goto Saida 
) else if %dispositivo% == 2 ( 
	goto DownLeg
) else if %dispositivo% == 3 ( 
	goto Joelho
) else if %dispositivo% == 4 ( 
	goto Nariz
) else (
	echo opcao invalida... Tente novamente
	pause >nul
	cls
	goto :ini
)

:All
:saida
	pushd "IR\Saida"
	call :processa
	popd
	pushd "Tratadas\Saida"
	call :processa
	popd
	if %dispositivo% neq 0 (
		goto FIM
	)
	
:DownLeg
	pushd "IR\DownLeg"
	call :processa
	popd
	pushd "Tratadas\DownLeg"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)

:Joelho
	pushd "IR\Joelho"
	call :processa
	popd
	pushd "Tratadas\Joelho"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)
	
:Nariz
	pushd "IR\Nariz"
	call :processa
	popd
	pushd "Tratadas\Nariz"
	call :processa
	popd
	
	if %dispositivo% neq 0 (
		goto FIM
	)

:FIM
	echo.
	echo Renomecao completa, favor conferir!!!!
	pause >nul
	goto :EOF


:processa
rem armazena o nomes na str
set /a cont=0
setlocal ENABLEDELAYEDEXPANSION
for %%s in (%str%) do (
	set nomecerto[!cont!]=%%~s_LD
	set /a cont+=1
	set nomecerto[!cont!]=%%~s_LE
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