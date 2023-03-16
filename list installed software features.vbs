' List Installed Software Features


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colFeatures = objWMIService.ExecQuery _
    ("Select * from Win32_SoftwareFeature")

For Each objFeature in colfeatures
    Wscript.Echo "Accesses: " & objFeature.Accesses
    Wscript.Echo "Attributes: " & objFeature.Attributes
    Wscript.Echo "Caption: " & objFeature.Caption
    Wscript.Echo "Description: " & objFeature.Description
    Wscript.Echo "Identifying Number: " & objFeature.IdentifyingNumber
    Wscript.Echo "Install Date: " & objFeature.InstallDate
    Wscript.Echo "Install State: " & objFeature.InstallState
    Wscript.Echo "Last Use: " & objFeature.LastUse
    Wscript.Echo "Name: " & objFeature.Name
    Wscript.Echo "Product Name: " & objFeature.ProductName
    Wscript.Echo "Vendor: " & objFeature.Vendor
    Wscript.Echo "Version: " & objFeature.Version
Next
