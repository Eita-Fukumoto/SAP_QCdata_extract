@echo off

echo P2R������������̂�Y�ꂸ��

rem QualityData�t�H���_�[��u�����p�X�ɏC�� (���L�̓f�X�N�g�b�v�ɒu�����ꍇ)
cd C:\Users\%username%\OneDrive - Bayer\Desktop\QualityData\Trend2

echo ������������Ă�������
set From=
set /p From= ������H:
set To= 
set /p To= ���܂ŁH:

if not exist "Files\%USERNAME%" mkdir "Files\%USERNAME%"
echo ver20241004 > Files\%USERNAME%\update_new.txt
echo who:%username% >> Files\%USERNAME%\update_new.txt
echo Start:%date% %time% >> Files\%USERNAME%\update_new.txt
echo From:%From% >> Files\%USERNAME%\update_new.txt
echo To:%To% >> Files\%USERNAME%\update_new.txt

echo Download of inspection plans
cscript /nologo scripts\Download_inspection_plans.vbs %To%

echo ���ԓ��̎g�p����i�ڂ̃f�[�^���o���J�n���܂�
cscript /nologo scripts\LotUpdate.vbs %From% %To% 
cscript /nologo scripts\Finish.vbs 

python scripts\writeTable.py || Pause 

echo qa33����f�[�^���_�E�����[�h���Ă��܂�
cscript /nologo scripts\qa33.vbs 
cscript /nologo scripts\Finish.vbs 

echo �������ʂ̃_�E�����[�h�����܂�

python scripts\Download_the_results.py

rem python scripts\CreateCSV.py
python scripts\CreateCSV_ver3.py

echo �������܂���

echo Finished %date% %time% >> Files\%USERNAME%\update_new.txt

type Log.txt  >> Files\%USERNAME%\update_new.txt 
type Files\%USERNAME%\update_new.txt > Log.txt 

msg * �X�V�I���܂���!

pause

