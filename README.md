# Wasy-Doc-Fill

## Installation

To install the required packages, run the following commands:

```sh
py -m pip install pandas
py -m pip install python-docx
py -m pip install ttkbootstrap
py -m pip install pyinstaller

or install all package just run this command:
pip install -r requirements.txt
```

For Windows don't forget to sett ENV

```sh
C:\Users\<username>\AppData\Local\Programs\Python\Python3x\Scripts
C:\Users\<username>\AppData\Local\Programs\Python\Python3x\Scripts
```

Make sure you have Python installed before running these commands.

Install .exe :

```sh
python -m PyInstaller --onefile --windowed --icon=icon/icon.ico EasyDocFill.py
```

copy the templates to /dist :

```sh
  Copy /templates to /dist
```

Final

```sh
Open the app from dir /dist EasyDocFil.exe
```
