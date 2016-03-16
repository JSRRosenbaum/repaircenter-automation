ECHO PRESS ENTER TO LAUNCH SUPER QC TEST written by Joseph Rosenbaum (JR)
ECHO INCLUDES WEBCAM TEST, SYSTEM PROPERTIES, DEVICE MANAGER, AUDIO DEVICES, AND WIFI TEST


pause


cd ".\Microsoft Silverlight\"
start Silverlight.exe

control /name Microsoft.System
control appwiz.cpl
control hdwwiz.cpl
control mmsys.cpl
netsh wlan connect name=WTX-Depot-New3

ECHO COMPLETE
pause
