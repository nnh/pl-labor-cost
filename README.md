# 概要
個人を特定する情報を可能な限りカットした状態で人件費を集計します。  
# 実行方法
- Zipファイルをダウンロードし、ダウンロードしたファイルを右クリックして「すべて展開」してください。 
![スクリーンショット 2021-08-12 16 40 27](https://user-images.githubusercontent.com/24307469/129157580-a9a88a5d-9f1d-4292-9caa-0502b8cbbad6.png)
- input/rawdataに入力ファイルを格納してください。
- Googleドキュメントの「部署メンバー一覧」を.xlsx形式でダウンロードし、input/extに格納してください。
- programs\pl-labor-cost.xlsmを開き、Sheet1の「実行」をクリックしてください。  
- ポップアップが表示されますので、表示に従って入力ファイルのパスワードと処理対象の年度（2019年度なら「2019」）を入力してください。  
- 「処理が終了しました」というポップアップが表示されたら処理完了です。outputフォルダの内容を確認してください。  
# プログラムの修正について
修正を行った後は、モジュールをエクスポートした後にGit bash等でtools\iconv.shを実行し、ファイルの文字コードを変換してからpushしてください。  
```
sh iconv.sh
```
# License
pl-labor-cost are licensed under the MIT license.  
Copyright © 2021, NHO Nagoya Medical Center and NPO-OSCR.  
