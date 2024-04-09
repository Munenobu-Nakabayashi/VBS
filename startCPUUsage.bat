rem //文字コード対応やで
rem chcp 65001

setlocal enabledelayedexpansion

pushd "C:\cpuusage"

type nul > .\cpuusage.log

:Loop
cscript "cpuRev2.vbs"

ping -n 30 localhost > nul

goto :Loop

popd

endlocal
exit 0