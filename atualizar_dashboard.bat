@echo off
setlocal

rem Executa a partir da pasta atual, mesmo com espaços no caminho
pushd "%~dp0"

echo Atualizando dashboard...
python "%~dp0generate_dashboard.py"
if %errorlevel% equ 0 (
  echo.
  echo Dashboard atualizado com sucesso.
) else (
  echo.
  echo Ocorreu um erro ao atualizar o dashboard.
)

echo.
pause

popd
endlocal
