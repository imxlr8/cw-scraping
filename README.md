# Cloudworks Scraping Tool 2022

## 使用したPythonのバージョン
    Python 3.10.4

## 使用したFlaskのバージョン
    Flask 2.2.2

Pythonのバージョンが古いとFlaskが動かないです

<br>

## ライブラリ関連
```
$ pip install Flask

$ pip install jinja2

$ pip install beautifulsoup4

$ pip install requests

$ pip install openpyxl
```

# 実行手順
・フォルダ『cw-scraping』内で `$ python main.py` を実行  

・http://127.0.0.1:5000　を開く  

・選択肢を絞り『人材を探す』ボタンをクリック、自動で取得と書き出しが始まります  

・同フォルダ内にExcelフォルダが書き出されていると思います  

※フォルダ名は「20XXMMDD選択した業種」となります。    
同日に書き出すと、他の絞り込み条件が違っても選択した業種が同じ場合上書きされてしまうので、  
忘れずに取り出してください_m_m_

<br>

## 取得する人数について  

・デフォルトではクラウドワークスの３ページ分となっております  
・main.pyの `PAGE` の値を変えることで取得するページ数を変更できます  
（1ページは最大30人です）

## Excelについて
シート内で適当にソートをかけると良い感じです  

## 構成
.  
┠━━ README.md  
┠━━ static  
┃　　　┗━━ style.css  
┠── templates  
┃　　　┠━━ data.html  
┃　　　┗━━ index.html  
┗── main.py  
