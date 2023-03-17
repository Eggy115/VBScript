' Save a Microsoft Word Document as an HTML File



Const wdFormatHTML = 8

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = _
    objWord.Documents.Open("c:\scripts\test.doc)
objDoc.SaveAs("C:\Scripts\test.htm", wdFormatHTML)
