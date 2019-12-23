@echo off
tasklist /fi "imagename eq report.exe" 2>NUL | find /i /n "report.exe" >NUL || start report.exe
exit
