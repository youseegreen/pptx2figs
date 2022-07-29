# pptx2figs
powerpoint中の図のpdf, pngを作成する

## How to use
1. スライド中で図にしたい領域を四角図形で囲む
2. 四角図形のテキストを<保存したい名前.pdf>とする
3. 全ての図にしたい領域に対して1, 2を繰り返す
4. ```python pptx2figs.py -i pptxファイルの名前``` を実行
5. pdfs, pngs, pptxsフォルダ内にそれぞれの図が作成される  

![](demo/demo.png)

## Requirements
- Windows OS (pywin32を使っているため)
- python (version 3以上なら大丈夫なはず)
- python-pptx
- pywin32

## Notion
- powerpointの数式が含まれている場合、正常に動作しません
- その他、特殊図形が使われている場合も動作しないと思います
- 作成されたpdfが正常かを十分に見るようにしてください
