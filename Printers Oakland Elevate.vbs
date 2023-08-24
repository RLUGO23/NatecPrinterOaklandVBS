Set UAC = CreateObject("Shell.Application")
UAC.ShellExecute "InstallPrinters.bat", "", "", "runas", 1