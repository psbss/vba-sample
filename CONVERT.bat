@echo off

:: --- 変数の宣言 ---
set base_file="sample.xlsm"

:: --- バッチファイル内容 ---

:: 実行時引数が -y の場合 :bool_exe_yes にスキップ
if {%1} == {-y} (goto :bool_exe_yes)

set /p bool_exe="VBAコードを結合しますか? (Y=Yes / N=NO)"
:: /i は大文字と小文字を制限せずに比較する
if /i {%bool_exe%} == {y} (goto :bool_exe_yes)
if /i {%bool_exe%} == {yes} (goto :bool_exe_yes)

:: "no" の場合処理終了
echo 処理を終了します
exit

:bool_exe_yes
:copy_file
    echo ベースファイルからコピーしています
    copy .\bin\base_xls\%base_file% .\bin\%base_file%

    if %ERRORLEVEL% equ 0 (
        echo コピーが成功しました
        echo コードの結合を開始します
        echo:

        :: vbacの実行
        cscript vbac.wsf combine
        echo コードの結合が終了しました
        echo:
    ) else (
        echo:
        echo 指定されたファイルが見つからないか、コピー先のExcelファイルが開いています。
        echo エンターキーを押すと再度実行します。
        pause
        goto :copy_file
    )
exit
