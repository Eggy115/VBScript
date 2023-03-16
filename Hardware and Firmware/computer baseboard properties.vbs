' List Computer Baseboard Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard")

For Each objItem in colItems
    For Each strOption in objItem.ConfigOptions
        Wscript.Echo "Configuration Option: " & strOption
    Next
    Wscript.Echo "Depth: " & objItem.Depth
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Height: " & objItem.Height
    Wscript.Echo "Hosting Board: " & objItem.HostingBoard
    Wscript.Echo "Hot Swappable: " & objItem.HotSwappable
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Other Identifying Information: " & _
        objItem.OtherIdentifyingInfo
    Wscript.Echo "Part Number: " & objItem.PartNumber
    Wscript.Echo "Powered-On: " & objItem.PoweredOn
    Wscript.Echo "Product: " & objItem.Product
    Wscript.Echo "Removable: " & objItem.Removable
    Wscript.Echo "Replaceable: " & objItem.Replaceable
    Wscript.Echo "Requirements Description: " & objItem.RequirementsDescription
    Wscript.Echo "Requires Daughterboard: " & objItem.RequiresDaughterBoard
    Wscript.Echo "Serial Number: " & objItem.SerialNumber
    Wscript.Echo "SKU: " & objItem.SKU
    Wscript.Echo "Slot Layout: " & objItem.SlotLayout
    Wscript.Echo "Special Requirements: " & objItem.SpecialRequirements
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Weight: " & objItem.Weight
    Wscript.Echo "Width: " & objItem.Width
Next
