# tcc

* 試験項目書生成ツール
* yaml形式で記載された試験項目をExcelに変換するツール。(プロトタイプ)

# コマンド

* tc-converter.py

```
Usage: tc-converter.py [OPTIONS] [FILES]...

Options:
  -f, --from-format TEXT  source file format.
  -t, --to-format TEXT    distination file format.
  -o, --output TEXT       excel file name
  --help                  Show this message and exit.
```

# 使い方

1. yaml形式で試験項目を作成
2. tc-converter.pyでexcelに変換
```
$ tc-converter.py -f yaml -t xlsx -o output.xlsx sample.yaml
```
