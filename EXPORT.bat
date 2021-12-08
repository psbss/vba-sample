@echo off

:: --- バッチファイル内容 ---

:: 実行時引数が -y の場合 :bool_exe_yes にスキップ
if {%1} == {-y} (goto :bool_exe_yes)

set /p selected="VBAコードを分離しますか? (Y=Yes / N=NO)"
:: /i は大文字と小文字を制限せずに比較する
if /i {%bool_exe%} == {y} (goto :bool_exe_yes)
if /i {%bool_exe%} == {yes} (goto :bool_exe_yes)

:: "no" の場合処理終了
echo 処理を終了します
exit

:bool_exe_yes
    pushd %0\..
    cscript vbac.wsf decombine
    echo コードの分離が終了しました
exit
