@echo off
setlocal
set "PYTHONIOENCODING=utf-8"
set "SCRIPT_DIR=%~dp0"
if defined PYTHON_EXE (
  "%PYTHON_EXE%" "%SCRIPT_DIR%query_project.py" %*
) else (
  python "%SCRIPT_DIR%query_project.py" %*
)
