@echo off
echo MD�t�@�C����Excel�ɕϊ�����f�������s���܂�
echo �T���v���t�@�C��: sample.md

python md_to_excel.py sample.md -o sample_output.xlsx -s "�v���W�F�N�g�v��"

echo.
echo �ϊ����������܂����Bsample_output.xlsx ���m�F���Ă��������B
echo.
pause
