@echo off
title %cd%
call conda activate
python run.py
exit