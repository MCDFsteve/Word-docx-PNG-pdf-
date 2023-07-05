@echo off
chcp 65001 > nul

setlocal

REM 获取脚本所在目录的绝对路径
for %%i in ("%~dp0.") do set "script_dir=%%~fi"
echo 安装docx2pdf
call pip install docx2pdf
echo 安装Pillow
call pip install Pillow
echo 安装fitz
call pip install fitz
echo 安装PyMuPDF
call pip install PyMuPDF
echo 安装frontend
call pip install frontend
REM 运行文档转png.py
call python "%script_dir%\文档转png.py"
echo 转换完成。作者DFSTEVE
pause
