
' Send Email from a Script


Set objEmail = CreateObject("CDO.Message")

objEmail.From = "monitor1@fabrikam.com"
objEmail.To = "admin1@fabrikam.com"
objEmail.Subject = "Atl-dc-01 down" 
objEmail.Textbody = "Atl-dc-01 is no longer accessible over the network."
objEmail.Send
