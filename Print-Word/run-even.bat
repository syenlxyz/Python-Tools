@echo off
for %%I in (%cd%) do title %%~nI
call conda activate
python run-even.py
pause