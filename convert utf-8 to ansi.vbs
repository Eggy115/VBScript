

' Convert UTF-8 file to ANSI

currentdir=Left(WScript.ScriptFullName,InStrRev(WScript.ScriptFullName,"\"))

source = currentdir & "source.txt"
dest = currentdir & "dest.txt"
charset= "Windows-1252"

Set stream=CreateObject("ADODB.Stream")
stream.Open
stream.Type = 1
stream.LoadFromFile source
stream.Type = 2
stream.Charset = "utf-8"
    
Dim fso
Set fso = CreateObject("Scripting.Filesystemobject")
    
Set f = fso.CreateTextFile(dest, True)
    
Do Until stream.EOS
   
  strLine = stream.ReadText(10000)
       
  Set output=CreateObject("ADODB.Stream")
  output.Open
  output.Type = 2
  output.Charset = charset
  output.WriteText strLine
       
  output.Position = 0
  str = output.ReadText(-1)

  f.Write str

Loop

f.Close
stream.Close  
