DIM C
Set C=CreateObject("Shell.Application")
C.ShellExecute "browser.exe","https://www.youtube.com","","",1
'browser.exe must be set to chrome.exe, msedge.exe, brave.exe, firefox.exe etc.
