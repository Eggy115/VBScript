do
Dim message, sapi
message=InputBox("What do you want me to say?","Eggy115's Voice Generator")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message
loop
