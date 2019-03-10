# 使い方
## 導入
* pythonをインストール(3.7で確認)
* 以下のコマンドでライブラリのインストールを行う
```
pip install numpy
pip install selenium
pip install pandas
pip install openpyxl
pip install xlrd
```
＊OSはwindows10似て確認
* ブラウザのインストール
    * __Chrome__ を使用するためFireFoxがインストールされていなければインストールすること。
        * ほかのブラウザでもdriverの指定を変えれば動くと思われるが未検証・・・
* webdriverの配置
    * [こちら](http://chromedriver.chromium.org/downloads)からダウンロード後解凍して`C:\webdriver` に格納する。配置場所を変えた場合はソースの該当箇所を変更する。
    FireFoxの場合は[こちら](https://github.com/mozilla/geckodriver/releases)から入手

## 使い方

1. `data/予約データ.xlsx`に各入力内容に応じて値を入れる。
    * 日付についてはランダムで入るように指定しているので固定したい場合は修正すること。
2. `seleniumTestSite1`を実行する。
3. 終了するのを待つ。`screenShot/reserve`に結果は出力される。

## ソースの説明

* 以前公開したクラスを継承して使用している。
* アラートダイアログに関しては基底クラスに定義していませんが定義するかも・・・
* エクセルなどでまとめてデータがある場合一気に自動入力できるような処理のサンプルとして作成。
    * 実際にはこの機能を応用して勤務表の入力やとあるシステムの業務支援として使っている。
* logの設定など見よう見まねでやっているのでベストプラクティスではないかもしれない
