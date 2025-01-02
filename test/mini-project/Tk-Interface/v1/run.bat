@echo off
if not defined minimized set minimized=1 && start /min run.bat && exit
title %cd%
call conda activate
python run.py
exit