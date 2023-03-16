' Send Email without Installing the SMTP Service


Set objEmail = CreateObject("CDO.Message")

objEmail.From = "admin1@fabrikam.com"
objEmail.To = "admin2@fabrikam.com"
objEmail.Subject = "Server down" 
objEmail.Textbody = "Server1 is no longer accessible over the network."
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
        "smarthost" 
objEmail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objEmail.Configuration.Fields.Update
objEmail.Send
