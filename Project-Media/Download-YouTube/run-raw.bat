@echo off
for %%I in (%cd%) do title %%~nI
call conda activate
python run-raw.py
pause