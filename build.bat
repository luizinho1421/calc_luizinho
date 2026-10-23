@echo off
chcp 65001 >nul
echo Compilando Calculadora (Created by LuizinhoSouzaK)...
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:winexe /out:"calc_luizinho.exe" /win32icon:"extracted\calc.ico" /optimize+ "CalculadoraLuizinho.cs"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [SUCESSO] Executavel calc_luizinho.exe gerado com sucesso!
) else (
    echo.
    echo [ERRO] Falha na compilacao.
)
pause
