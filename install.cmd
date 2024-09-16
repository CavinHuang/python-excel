@echo off
chcp 65001
pyinstaller --onefile --noconsole --noupx ./src/ui.py --name "Excel自动化处理程序-另一个版本"