call conda activate base

SET CurrentDir="%~dp0"
cd /D %CurrentDir%

python fin_plan2.py %1

pause