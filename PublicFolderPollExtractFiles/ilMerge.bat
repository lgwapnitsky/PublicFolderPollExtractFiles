ECHO parameter1=$1
ECHO parameter2=$2
cd %1
"c:\program files (x86)\Microsoft\ILMerge\ILMerge.exe" /out:%2.dll Microsoft.Exchange.WebServices.dll Powerargs.dll /log:log.txt