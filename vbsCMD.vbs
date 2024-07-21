Set objShell = CreateObject("Shell.Application")


objShell.ShellExecute "cmd.exe", "taskkill /im excel.exe /f" , "", "runas", 10

