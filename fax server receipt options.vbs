

' List Fax Server Receipt Options


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objReceiptOptions = objFaxServer.ReceiptOptions

Wscript.Echo "Allowed receipts: " & _
    objReceiptOptions.AllowedReceipts
Wscript.Echo "Authentication type: " & _
    objReceiptOptions.AuthenticationType
Wscript.Echo "SMTP password: " & objReceiptOptions.SMTPPassword
Wscript.Echo "SMTP port: " & objReceiptOptions.SMTPPort
Wscript.Echo "SMTP sender: " & objReceiptOptions.SMTPSender
Wscript.Echo "SMTP server: " & objReceiptOptions.SMTPServer
Wscript.Echo "SMTP user: " & objReceiptOptions.SMTPUser
Wscript.Echo "Use for inbound routing: " & _
    objReceiptOptions.UseForInboundRouting
