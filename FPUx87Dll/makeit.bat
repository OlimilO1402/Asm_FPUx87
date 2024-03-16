@echo off
if exist C:\Masm32Projects\FPUx87\FPUx87.obj del C:\Masm32Projects\FPUx87\FPUx87.obj
if exist C:\Masm32Projects\FPUx87\FPUx87.dll del C:\Masm32Projects\FPUx87\FPUx87.dll
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_FPUx87\FPUx87Dll\FPUx87.asm" "C:\Masm32Projects\FPUx87\FPUx87.asm"
copy "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_FPUx87\FPUx87Dll\FPUx87.def" "C:\Masm32Projects\FPUx87\FPUx87.def"
cd C:\Masm32Projects\FPUx87\
C:\Masm32Projects\FPUx87\makeit.bat
rem C:\masm32\bin\ml /c /coff C:\Masm32Projects\FPUx87\FPUx87.asm
rem C:\masm32\bin\Link /SUBSYSTEM:WINDOWS /DLL /DEF:C:\Masm32Projects\FPUx87\FPUx87.def C:\Masm32Projects\FPUx87\FPUx87.obj 
copy "C:\Masm32Projects\FPUx87\FPUx87.dll" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_FPUx87\FPUx87Dll\FPUx87.dll"
copy "C:\Masm32Projects\FPUx87\FPUx87.lib" "\\SOLS_DS\Daten\GitHubRepos\VB\Asm_FPUx87\FPUx87Dll\FPUx87.lib"
del C:\Masm32Projects\FPUx87\FPUx87.obj
del C:\Masm32Projects\FPUx87\FPUx87.exp
dir C:\Masm32Projects\FPUx87\FPUx87.*
pause