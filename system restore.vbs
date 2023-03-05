localdate=date()
set SRP = getobject("winmgmts:\\.\root\default:Systemrestore")
CSRP = SRP.createrestorepoint ("System Start-up - " & localdate, 0, 100)
