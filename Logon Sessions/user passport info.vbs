

' List User Passport Information


Set objUser = CreateObject("UserAccounts.PassportManager")
Wscript.Echo "Current Passport: " & objUser.CurrentPassport
Wscript.Echo "Member services URL: " & objUser.MemberServicesURL
