' List Installed Administrative Tools


Const ADMINISTRATIVE_TOOLS = &H2f&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(ADMINISTRATIVE_TOOLS) 
Set objTools = objFolder.Items

For i = 0 to objTools.Count - 1
    Wscript.Echo objTools.Item(i)
Next

