python CornDataProcessor.py
copy "E:\Desktop\PyCode\webdata.xlsx" "E:\Desktop\fhzbWeb\report\webdata.xlsx"
copy "E:\Desktop\workstation\database\DeepProcessOprationAndProfit.xlsx" "E:\Desktop\fhzbWeb\report\DeepProcessOprationAndProfit.xlsx"
copy "E:\Desktop\workstation\database\CornTempRes.xlsx" "E:\Desktop\fhzbWeb\report\CornTempRes.xlsx"

pscp -pw Dean0129 "E:\Desktop\fhzbWeb\report\webdata.xlsx" ubuntu@111.230.170.168:/home/ubuntu/fhzbWeb/report/
pscp -pw Dean0129 "E:\Desktop\fhzbWeb\report\DeepProcessOprationAndProfit.xlsx" ubuntu@111.230.170.168:/home/ubuntu/fhzbWeb/report/
pscp -pw Dean0129 "E:\Desktop\fhzbWeb\report\CornTempRes.xlsx" ubuntu@111.230.170.168:/home/ubuntu/fhzbWeb/report/
pause()