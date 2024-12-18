call conda activate base

SET CurrentDir="%~dp0"
cd /D %CurrentDir%

python fin_plan4.py -f %1

pause