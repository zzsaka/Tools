@echo off
echo MDファイルをExcelに変換するデモを実行します
echo サンプルファイル: sample.md

python md_to_excel.py sample.md -o sample_output.xlsx -s "プロジェクト計画"

echo.
echo 変換が完了しました。sample_output.xlsx を確認してください。
echo.
pause
