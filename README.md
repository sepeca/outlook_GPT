1) Base interpreter Python 3.x.
2) Execute in console
pip install --upgrade pip
pip install -r requirements.txt

3) For creating .exe

.\.venv\Scripts\python.exe -m PyInstaller --onefile --noconsole `
   --name OutlookGPT_GUI `
   --icon app.ico `
   --collect-submodules win32com `
   --hidden-import win32timezone `
   --add-data ".env.example;." `
   gui_env.py

If you want to try without creating .exe, create .env as .env.example, execute gui_env.py
For proper result in your excel example you need headers be on 2-nd row and every raw from 3 empty.
Headers with names:
1) Název klienta*/Název klienta/NazevKlienta
2) Příjmení*/Příjmení/Prijmeni
3) Jméno/Jmeno
4) Titul před/TitulPred
5) Titul za/TitulZa
6) Funkce
7) Tel 1/Tel1/Telefon
8) E-mail/Email
9) WWW
10) PoznamkaKOsobe/Poznámka k osobě

As soon as surname and name (Prijmeni jmeno) are the same in 2 different rows, this rows will be filled with pink color.
Its done to call user`s attention to decide if its same person and edit.