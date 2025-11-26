@echo off
rem 一键使用本项目虚拟环境启动 GUI（避免系统 Python 环境不一致）
setlocal
cd /d %~dp0
".\.venv\Scripts\python.exe" weibo_hot_gui.py