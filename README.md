# GoogleSheets-SDRTouch-Exporter

## これは何？
Google スプレッドシートで周波数をブックマークし、
Android アプリの SDR Touch にまとめてエクスポートするための GAS スクリプトです。

## 使い方
次のリンクをご自身の Google Drive にコピーしていただき、上部メニューの [SDR Tools] から [Export Frequencies] を
選択すると SDR Touch 向けの xml が出力されます。
[SDRTouch Radio Frequencies](https://docs.google.com/spreadsheets/d/1m8zwJe1khnVhV9HNBPGA2ebZTcewB-h_3vrwBfDg4pI/edit?usp=sharing)

これを SDR Touch 上からインポートすることで、まとめてブックマークすることが可能です。

## 謝辞
このスクリプトは次の記事をベース/参考に作成いたしました。
[Google Sheets SDR Touch Exporter - Version 2.0](http://www.blogbyben.com/2018/12/google-sheets-sdr-touch-exporter.html)

また、周波数計算にて下記のプロジェクトをライブラリとして使用しています。
[hiroshi-manabe/JSDecimal](https://github.com/hiroshi-manabe/JSDecimal)