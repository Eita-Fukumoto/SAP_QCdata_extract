@echo off

echo P2Rをたちあげるのを忘れずに

rem QualityDataフォルダーを置いたパスに修正 (下記はデスクトップに置いた場合)
cd C:\Users\%username%\OneDrive - Bayer\Desktop\QualityData\Trend2

echo 判定日を教えてください
set From=
set /p From= いつから？:
set To= 
set /p To= いつまで？:

if not exist "Files\%USERNAME%" mkdir "Files\%USERNAME%"
echo ver20241004 > Files\%USERNAME%\update_new.txt
echo who:%username% >> Files\%USERNAME%\update_new.txt
echo Start:%date% %time% >> Files\%USERNAME%\update_new.txt
echo From:%From% >> Files\%USERNAME%\update_new.txt
echo To:%To% >> Files\%USERNAME%\update_new.txt

echo Download of inspection plans
cscript /nologo scripts\Download_inspection_plans.vbs %To%

echo 期間内の使用決定品目のデータ抽出を開始します
cscript /nologo scripts\LotUpdate.vbs %From% %To% 
cscript /nologo scripts\Finish.vbs 

python scripts\writeTable.py || Pause 

echo qa33からデータをダウンロードしています
cscript /nologo scripts\qa33.vbs 
cscript /nologo scripts\Finish.vbs 

echo 試験結果のダウンロードをします

python scripts\Download_the_results.py

rem python scripts\CreateCSV.py
python scripts\CreateCSV_ver3.py

echo 完了しました

echo Finished %date% %time% >> Files\%USERNAME%\update_new.txt

type Log.txt  >> Files\%USERNAME%\update_new.txt 
type Files\%USERNAME%\update_new.txt > Log.txt 

msg * 更新終わりました!

pause

