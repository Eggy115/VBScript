' List the Attributes of the organizationalUnit Class


Set objOrganizationalUnitClass = _
    GetObject("LDAP://schema/organizationalUnit")

Set objSchemaClass = GetObject(objOrganizationalUnitClass.Parent)
 
i = 0
WScript.Echo "Mandatory attributes:"

For Each strAttribute in objOrganizationalUnitClass.MandatoryProperties
    i= i + 1
    WScript.Echo i & vbTab & strAttribute
    Set objAttribute = objSchemaClass.GetObject("Property",  strAttribute)
    WScript.Echo " (Syntax: " & objAttribute.Syntax & ")"
    If objAttribute.MultiValued Then
        WScript.Echo " Multivalued"
    Else
        WScript.Echo " Single-valued"
    End If
Next
 
WScript.Echo VbCrLf & "Optional attributes:"
For Each strAttribute in objOrganizationalUnitClass.OptionalProperties
    i= i + 1
    WScript.StdOut.Write i & vbTab & strAttribute
    Set objAttribute = objSchemaClass.GetObject("Property",  strAttribute)
    Wscript.Echo " [Syntax: " & objAttribute.Syntax & "]"
    If objAttribute.MultiValued Then
        WScript.Echo " Multivalued"
    Else
        WScript.Echo " Single-valued"
    End If
Next
