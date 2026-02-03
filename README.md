# Python + Tkinter десктоп приложение, което:

* качва 2 Excel файла: Поръчка (.xls/.xlsx) и Цени (.xls/.xlsx)

* чете нужните колони (намира ги по заглавия)

* слива по “Име на артикул”

* избира Ед. цена според тираж (количество) от колоните 1000/2000/3000…

* взима най-малкия тираж ≥ бройки, ако няма — взима най-големия наличен

* смята Сума = Ед. цена * Бройки

* показва резултата в таблица

* записва готов .xlsx (Save As)

macOS / Linux
```
python3 -m venv .venv
source .venv/bin/activate
pip install -U pip
pip install pandas openpyxl xlrd==2.0.1 pyinstaller
```

Важно: build-ът трябва да се прави на същата OS, за която е executable-ът.

macOS: .app
```
pyinstaller --windowed --onefile --name "MergeOrders" merge_excel_app.py
```
Windows: .exe (без конзола)

В папката, където е merge_excel_app.py:
```
pyinstaller --noconsole --onefile --name "MergeOrders" merge_excel_app.py
```

```
python3 -m PyInstaller --noconsole --onefile --name "MergeOrders" --collect-all pandas --collect-all openpyxl --collect-all xlrd merge_excel_app.py
```