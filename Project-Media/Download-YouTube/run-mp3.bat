@echo off
for %%I in (%cd%) do title %%~nI
call conda activate
python run-mp3.py
pause