
' List RSOP Application Management Policy Settings


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery _
    ("Select * from RSOP_ApplicationManagementPolicySetting")

For Each objItem in colItems  
    Wscript.Echo "Allow X86 on IA64: " & objItem.AllowX86OnIA64
    Wscript.Echo "Application ID: " & objItem.ApplicationId
    Wscript.Echo "Apply Cause: " & objItem.ApplyCause
    Wscript.Echo "Assignment Type: " & objItem.AssignmentType
    Wscript.Echo "Categories: " & objItem.Categories
    Wscript.Echo "Demand Installable: " & objItem.DemandInstallable
    Wscript.Echo "Deployment Last Modify Time: " & _
        objItem.DeploymentLastModifyTime
    Wscript.Echo "Deployment Type: " & objItem.DeploymentType
    Wscript.Echo "Display in Add/Remove Programs: " & objItem.DisplayInARP
    Wscript.Echo "Eligibility: " & objItem.Eligibility
    Wscript.Echo "Entry Type: " & objItem.EntryType
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "Ignore Language: " & objItem.IgnoreLanguage
    Wscript.Echo "Installation UI: " & objItem.InstallationUI
    Wscript.Echo "Language ID: " & objItem.LanguageId
    Wscript.Echo "Language Match: " & objItem.LanguageMatch
    Wscript.Echo "Loss of Scope Action: " & objItem.LossOfScopeAction
    For Each strArchitecture in objItem.MachineArchitectures
        Wscript.Echo "Machine Architecture: " & strArchitecture
    Next
    Wscript.Echo "On-demand CLSID: " & objItem.OnDemandClsid
    Wscript.Echo "On-demand File Extension: " & objItem.OnDemandFileExtension
    Wscript.Echo "On-demand ProgID: " & objItem.OnDemandProgId
    Wscript.Echo "Package Location: " & objItem.PackageLocation
    Wscript.Echo "Package Type: " & objItem.PackageType
    Wscript.Echo "Precedence: " & objItem.Precedence
    Wscript.Echo "Precedence Reason: " & objItem.PrecedenceReason
    Wscript.Echo "Product ID: " & objItem.ProductId
    Wscript.Echo "Publisher: " & objItem.Publisher
    Wscript.Echo "Redeploy Count: " & objItem.RedeployCount
    Wscript.Echo "Removal Cause: " & objItem.RemovalCause
    Wscript.Echo "Removal Type: " & objItem.RemovalType
    Wscript.Echo "Removing Application: " & objItem.RemovingApplication
    Wscript.Echo "Replaceable Applications: " & objItem.ReplaceableApplications
    Wscript.Echo "Script File: " & objItem.ScriptFile
    Wscript.Echo "Support URL: " & objItem.SupportURL
    Wscript.Echo "Transforms: " & objItem.Transforms
    Wscript.Echo "Uninstall Unmanaged: " & objItem.UninstallUnmanaged
    Wscript.Echo "Upgradeable Applications: " & objItem.UpgradeableApplications
    Wscript.Echo "Upgrade Settings Mandatory: " & _
        objItem.UpgradeSettingsMandatory
    Wscript.Echo "Version Number (High): " & objItem.VersionNumberHi
    Wscript.Echo "Version Number (Low): " & objItem.VersionNumberLo
    Wscript.Echo
Next
