:: 1行目：文字コード指定
:: 2行目：anaconda(3)/Scripts/activate.batがあるフォルダを指定
:: 3行目：pywin32, python-pptxがある環境にアクティベートする
:: 4行目：指定した環境内のpython.exeを実行する
:: 上のコメントは動作成功したら消しといてください
chcp 65001
call C:/anaconda/Scripts/activate.bat
call activate eh
C:/anaconda/envs/eh/python.exe pptx2figs.py