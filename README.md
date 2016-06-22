AECalc
======
An Excel application with VBA that automatically calculates and shows CTCAE grades based on patients' age, gender and lab data.

[![MIT licensed][shield-license]](#)

Table of Contents
-----------------
  * [Requirements](#requirements)
  * [Usage](#usage)
  * [Contributing](#contributing)
  * [License](#license)

Requirements
------------
AECalc requires the following to run:
  * Microsoft Excel

Usage
-----
  * [AECalc.xlsm][AECalc]をダウンロードして中に記載されている使用方法を参照下さい。
  * Download [AECalc.xlsm][AECalc] and read the instruction in it.

Contributing
------------
  * [nkf][nkf]
    - ソースコード中の日本語のコード変換を自動的に行うため（shift-JIS←→UTF-8）上記のリンクからファイルをダウンロードし、ソースコードのフォルダの配下に `bin` というフォルダを作成し、nkf32.exe を配置してください。使うだけなら不要です。
    - To use Japanese characters in the program, you'll need to download nkf32.exe and put it into 'bin' folder where your source codes are located in order to avoid garbled on Github . You don't need nkf to run AECalc.

  * 開発手順
    1. git で最新のバージョンのソースコードをプルします。
    2. import.bat で AECalc.xlsm に最新のマクロを反映します。
    3. AECalc.xlsm で開発、編集し、保存します。
    4. export.bat バッチを実行します。
    5. git でファイルをコミット、プッシュします。

  * Develpment procedure
    1. git pull latest version
    2. reflect latest macro by running import.bat
    3. develop on the file "AECalc.xlsm"
    4. run export.bat
    5. git commit and push the files

License
-------
AECalc is licensed under the [MIT](#) license.  
Copyright &copy; 2016, NHO Nagoya Medical Center and NPO-OSCR

[nkf]: http://www.vector.co.jp/soft/dl/win95/util/se295331.html
[shield-license]: https://img.shields.io/badge/license-MIT-blue.svg
[AECalc]: https://github.com/nnh/AECalc/blob/master/AECalc.xlsm
