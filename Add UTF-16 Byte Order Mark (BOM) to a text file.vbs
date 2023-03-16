Set toolkit = CreateObject("Vbsedit.Toolkit")

'For UTF-16 Big Endian - FE FF
toolkit.AddBOM "C:\my_utf8_file.txt",1 

'For UTF-16 Little Endian - FF FE
toolkit.AddBOM "C:\my_utf8_file.txt",2 
