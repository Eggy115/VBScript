' Draw the Spiral of Theodorus as a PNG image
' https://en.wikipedia.org/wiki/Theodorus_of_Cyrene

angle=0
adj=20
opp=20

Dim arr(1000,1)

Dim img
Set img = CreateObject("Vbsedit.ImageProcessor")
img.Color="red"
img.BrushColor="lightgreen"
  
currentdir=Left( WScript.ScriptFullName, InStrRev( WScript.ScriptFullName,"\"))

img.Create 1000,1000,"White"

x1=adj
y1=0

count=500

arr(0,0)=x1
arr(0,1)=y1
  
For i=1 to count
  hyp=Sqr(adj*adj + opp*opp)
  angle = Atn(opp/adj)
  angletotal = angletotal + angle
  x=Cos(angletotal)*hyp
  y=Sin(angletotal)*hyp
  arr(i,0)=x
  arr(i,1)=y
  adj=hyp
Next

For i=count-1 to 0 step -1
  x=arr(i+1,0)
  y=arr(i+1,1)

  x1=arr(i,0)
  y1=arr(i,1)
  

  img.FillPolygon 500,500,500+x,500-y,500+x1,500-y1

  img.DrawPolygon 500,500,500+x,500-y,500+x1,500-y1

Next


img.Save currentdir & "spiral.png"
Set shell = WScript.CreateObject("Wscript.Shell")
shell.Run currentdir & "spiral.png",1,False
