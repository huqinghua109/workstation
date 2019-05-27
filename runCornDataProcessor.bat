cd /d %~dp0

python CornDataProcessor.py

copy .\webdata.xlsx ..\fhzbWeb\report\webdata.xlsx
copy .\database\DeepProcessOprationAndProfit.xlsx ..\fhzbWeb\report\DeepProcessOprationAndProfit.xlsx
copy .\database\CornTempRes.xlsx ..\fhzbWeb\report\CornTempRes.xlsx

pscp -pw Dean0129 .\webdata.xlsx ubuntu@111.230.170.168:/home/ubuntu/fhzbWeb/report/
pscp -pw Dean0129 .\database\DeepProcessOprationAndProfit.xlsx ubuntu@111.230.170.168:/home/ubuntu/fhzbWeb/report/
pscp -pw Dean0129 .\database\CornTempRes.xlsx ubuntu@111.230.170.168:/home/ubuntu/fhzbWeb/report/
pause()