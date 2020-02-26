@echo off
net stop BITS
ipconfig /flushdns
ren "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr0.dat" qmgr0.dat.old
ren "%ALLUSERSPROFILE%\Application Data\Microsoft\Network\Downloader\qmgr1.dat" qmgr1.dat.old
net start BITS