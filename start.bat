@echo off
cd /d "%~dp0"
title Listing Runner
python poller.py %*
pause
