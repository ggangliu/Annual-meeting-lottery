import os

os.system("pyinstaller -F Annual_awards_main.py")
os.system("copy 9224.wav dist\\")
os.system("copy bg_1366x768.png dist\\")
os.system("copy bg.png dist\\")
os.system("copy gc_cz.png dist\\")
os.system("copy name_file.xlsx dist\\")
os.system("copy simhei.ttf dist\\")