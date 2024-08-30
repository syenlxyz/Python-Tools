@echo off
title %cd%
call conda activate
python run-duplex.py
pause