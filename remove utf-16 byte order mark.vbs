' Remove UTF-16 Byte Order Mark (BOM) from a text file

Set toolkit = CreateObject("Vbsedit.Toolkit")
toolkit.RemoveBOM "C:\my_utf8_file.txt"
