import os

os.system("pyinstaller -F Annual_awards_main.py")
os.system("copy 9224.wav dist\\")
os.system("copy bg.png dist\\")
os.system("copy cz.png dist\\")
os.system("copy name_file.xlsx dist\\")