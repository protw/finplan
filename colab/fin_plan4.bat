call conda activate base

SET CurrentDir="%~dp0"
cd /D %CurrentDir%/finplan

python fin_plan4.py %1

pause