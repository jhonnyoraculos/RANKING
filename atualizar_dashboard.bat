@echo off
setlocal EnableExtensions

rem Executa a partir da pasta atual, mesmo com espacos no caminho
pushd "%~dp0"
if %errorlevel% neq 0 (
  echo Nao foi possivel acessar a pasta do script.
  goto :FIM
)

echo Atualizando dashboard...
python "%~dp0generate_dashboard.py"
if %errorlevel% neq 0 (
  echo.
  echo Ocorreu um erro ao atualizar o dashboard.
  goto :FIM
)

echo.
echo Dashboard atualizado com sucesso.
if exist "%~dp0index.html" if exist "%~dp0dashboard.html" (
  del /q "%~dp0dashboard.html"
  echo Arquivo legado dashboard.html removido. Principal: index.html
)
echo.
echo Sincronizando com Git...

git rev-parse --is-inside-work-tree >nul 2>&1
if %errorlevel% neq 0 (
  echo Repositorio Git nao encontrado nesta pasta.
  goto :FIM
)

git add -A

set "HAS_CHANGES="
for /f %%i in ('git diff --cached --name-only') do (
  set "HAS_CHANGES=1"
)

if not defined HAS_CHANGES (
  echo Nenhuma alteracao para enviar ao Git.
  goto :FIM
)

git commit -m "chore: atualizar dashboard automatico"
if %errorlevel% neq 0 (
  echo Falha ao criar commit.
  goto :FIM
)

git push
if %errorlevel% equ 0 (
  echo Push concluido com sucesso.
) else (
  echo Falha ao enviar para o remoto.
)

:FIM
echo.
pause
popd
endlocal
