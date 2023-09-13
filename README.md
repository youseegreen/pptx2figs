# pptx2figs
powerpoint中の図のpdf, pngを作成する

## How to use
1. スライド中で図にしたい領域を四角図形で囲む
2. 四角図形のテキストを<保存したい名前.pdf>とする
3. 全ての図にしたい領域に対して1, 2を繰り返す
4. ```python pptx2figs.py -i pptxファイルの名前``` を実行
5. pdfs, pngs, pptxsフォルダ内にそれぞれの図が作成される  

![](demo/demo.png)

## ボタン1クリックで実行するために...
1. ```pptx2figs.bat``` をテキストエディタで開き、anacondaのパスや環境名を自分の環境に合わせる。
2. inputsというフォルダを作り、その中にpptxファイルを配置する
3. ```pptx2figs.bat``` をダブルクリック
4. pdfs, pngs, pptxsフォルダ内にそれぞれの図が作成される  


## Requirements
- Windows OS (pywin32を使っているため)
- python (version 3以上なら大丈夫なはず)
- python-pptx
- pywin32

## Notion
- powerpointの数式が含まれている場合、正常に動作しません  
（数式の移動が発生しないだけのため、図の左上をページ左上に合わせれば対処できます。）
- その他、特殊図形が使われている場合も動作しないと思います
- 作成されたpdfが正常かを十分に見るようにしてください
