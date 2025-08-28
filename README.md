

.\.venv\Scripts\python.exe -m PyInstaller --onefile --noconsole `
>>   --name OutlookGPT_GUI `
>>   --icon app.ico `
>>   --collect-submodules win32com `
>>   --hidden-import win32timezone `
>>   --add-data ".env.example;." `
>>   gui_env.py
