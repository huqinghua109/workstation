cd /d %~dp0

python CornDataProcessor.py

copy .\webdata.xlsx ..\fhzbWeb\report\webdata.xlsx
copy .\database\DeepProcessOprationAndProfit.xlsx ..\fhzbWeb\report\DeepProcessOprationAndProfit.xlsx
copy .\database\CornTempRes.xlsx ..\fhzbWeb\report\CornTempRes.xlsx

pause()