Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\user\Desktop\클로드 코워크"
WshShell.Run """C:\Users\user\AppData\Local\Programs\Python\Python311\python.exe"" main.py", 0, False
