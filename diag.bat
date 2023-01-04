@echo off 
::Diagnostic file created @ 18:07:11.94 on Fri 12/16/2016 
copy nul > diagpass.txt 
PING 127.0.0.1 -n 1 -w 10 >NUL 
exit 
