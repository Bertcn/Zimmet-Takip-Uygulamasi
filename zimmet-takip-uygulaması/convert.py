from PyQt5 import uic

with open("panel.py","w",encoding="utf-8") as fout:
    uic.compileUi("tablo.ui",fout)