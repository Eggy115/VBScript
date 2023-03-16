' Draw the Spiral of Theodorus in Scalable Vector Graphics (SVG) 
' https://en.wikipedia.org/wiki/Theodorus_of_Cyrene

angle=0
adj=20
opp=20

Set fso = CreateObject ("Scripting.Filesystemobject")

currentdir=Left( WScript.ScriptFullName, InStrRev(WScript.ScriptFullName,"\"))

Set svg = fso.CreateTextFile(currentdir & "spiral.htm",True)
svg.WriteLine "<!DOCTYPE html>"
svg.WriteLine "<html>"
svg.WriteLine "<body>"
svg.WriteLine "<svg xmlns=""http://www.w3.org/2000/svg"" width=""800"" height=""800"" viewBox=""0 0 1000 1000"">"

lines=""

x1=adj
y1=0
For i=0 to 500
  hyp=Sqr(adj*adj + opp*opp)
  angle = Atn(opp/adj)
  angletotal = angletotal + angle
  x=Cos(angletotal)*hyp
  y=Sin(angletotal)*hyp
  adj=hyp
  
  newline = "<polygon points=""500 500, " & F(500+x)  & " " & F(500-y) & ", " & F(500+x1) & " " & F(500-y1) & """ style=""fill:lime;stroke:purple;stroke-width:1""/>"
  lines = newline & vbCrLf & lines
  x1=x
  y1=y
Next

svg.Write lines

svg.WriteLine "</svg></body></html>"

Set shell = CreateObject("Wscript.Shell")
shell.Run currentdir & "spiral.htm",1,False

Function F(z)
  F=Replace(z,",",".")
End Function
