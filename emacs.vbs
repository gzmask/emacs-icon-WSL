Set WshShell = CreateObject("WScript.Shell")
WshShell.Run """C:\Program Files\VcXsrv\vcxsrv.exe""  :0 -ac -terminate -lesspointer -multiwindow -clipboard -wgl"
WshShell.Run "C:\Windows\System32\bash.exe -ic ""/usr/local/bin/emacs-25.1 --insecure""", 0
Set WshShell = Nothing