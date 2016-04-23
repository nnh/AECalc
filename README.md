# AECalc

## requirements

### [nkf](http://www.vector.co.jp/soft/dl/win95/util/se295331.html)
- To use Japanese characters in the program, you'll need to download nkf32.exe and put it into 'bin' folder where your source codes are located.
- 上記のリンクからファイルをダウンロードし、ソースコードのフォルダの配下に `bin` というフォルダを作成し、nkf32.exe を配置してください

## 開発手順

1. git で最新のバージョンのソースコードをプルします。
1. import.bat で AECalc.xlsm に最新のマクロを反映します。
1. AECalc.xlsm で開発、編集し、保存します。
1. export.bat バッチを実行します。
1. git でファイルをコミット、プッシュします。

## Develpment procedure
1. git pull latest version
2. reflect latest macro by running import.bat
3. develop on the file "AECalc.xlsm"
4. run export.bat
5. git commit and push the files
