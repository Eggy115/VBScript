' Read an Active Directory Security Descriptor


Const ADS_RIGHT_DELETE = &H10000
Const ADS_RIGHT_READ_CONTROL = &H20000
Const ADS_RIGHT_WRITE_DAC = &H40000
Const ADS_RIGHT_OWNER = &H80000
Const ADS_RIGHT_SYNCHRONIZE = &H100000
Const ADS_RIGHT_ACCESS_SYSTEM_SECURITY = &H1000000
Const ADS_RIGHT_GENERIC_READ = &H80000000
Const ADS_RIGHT_GENERIC_WRITE = &H40000000
Const ADS_RIGHT_GENERIC_EXECUTE = &H20000000
Const ADS_RIGHT_GENERIC_ALL = &H10000000
Const ADS_RIGHT_DS_CREATE_CHILD = &H1
Const ADS_RIGHT_DS_DELETE_CHILD = &H2
Const ADS_RIGHT_ACTRL_DS_LIST = &H4
Const ADS_RIGHT_DS_SELF = &H8
Const ADS_RIGHT_DS_READ_PROP = &H10
Const ADS_RIGHT_DS_WRITE_PROP = &H20
Const ADS_RIGHT_DS_DELETE_TREE = &H40
Const ADS_RIGHT_DS_LIST_OBJECT = &H80
Const ADS_RIGHT_DS_CONTROL_ACCESS = &H100
Const ADS_ACETYPE_ACCESS_ALLOWED = &H0
Const ADS_ACETYPE_ACCESS_DENIED = &H1
Const ADS_ACETYPE_SYSTEM_AUDIT = &H2
Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &H5
Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const ADS_ACETYPE_SYSTEM_AUDIT_OBJECT = &H7

Set objSdUtil = GetObject("LDAP://OU=Finance,DC=fabrikam,DC=Com")
Set objSD = objSdUtil.Get("ntSecurityDescriptor")
Set objDACL = objSD.DiscretionaryACL

For Each objACE in objDACL
    Wscript.Echo "Trustee: " & objACE.Trustee

    If objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED Then
        Wscript.Echo "Ace Type: Access Allowed"
    ElseIf objACE.AceType = ADS_ACETYPE_ACCESS_DENIED Then
        Wscript.Echo "Ace Type: Access Denied"
    ElseIf objACE.AceType = ADS_ACETYPE_SYSTEM_AUDIT Then
        Wscript.Echo "Ace Type: System Audit "
    ElseIf objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT Then
        Wscript.Echo "Ace Type: Access Allowed"
    ElseIf objACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT Then
        Wscript.Echo "Ace Type: Access Denied"
    ElseIf objACE.AceType = ADS_ACETYPE_SYSTEM_AUDIT_OBJECT Then
        Wscript.Echo "Ace Type: System Audit"
    Else
        Wscript.Echo "Ace type could not be determined."
    End If

    If objACE.AccessMask And ADS_RIGHT_DELETE Then
        Wscript.Echo vbTab & "Delete"
    End If

    If objACE.AccessMask And ADS_RIGHT_READ_CONTROL Then
        Wscript.Echo vbTab & "Read from the security descriptor (not including the SACL)"
    End If

    If objACE.AccessMask And ADS_RIGHT_WRITE_DAC Then
        Wscript.Echo vbTab & "Modify the DACL"
    End If

    If objACE.AccessMask And ADS_RIGHT_OWNER Then
        Wscript.Echo vbTab & "Take ownership"
    End If

    If objACE.AccessMask And ADS_RIGHT_SYNCHRONIZE Then
        Wscript.Echo vbTab & "Use the object for synchronization"
    End If

    If objACE.AccessMask And RIGHT_ACCESS_SYSTEM_SECURITY Then
        Wscript.Echo vbTab & "Get or set the SACL"
    End If

    If objACE.AccessMask And ADS_RIGHT_GENERIC_READ Then
        Wscript.Echo vbTab & "Read permissions and properties"
    End If

    If objACE.AccessMask And ADS_RIGHT_GENERIC_WRITE Then
        Wscript.Echo vbTab & "Write permissions and properties"
    End If

    If objACE.AccessMask And ADS_RIGHT_GENERIC_EXECUTE Then
        Wscript.Echo vbTab & "Read permissions on and list the contents of the container"
    End If

    If objACE.AccessMask And ADS_RIGHT_GENERIC_ALL Then
        Wscript.Echo vbTab & "Create or delete child objects, delete a subtree, read and write " & _
            "properties, examine child objects and the object itself, add and remove the " & _
                "object from the directory, and read or write with an extended right"
    End If
  
    If objACE.AccessMask And ADS_RIGHT_DS_CREATE_CHILD Then
        Wscript.Echo vbTab & "Create child objects"
    End If

    If objACE.AccessMask And ADS_RIGHT_DS_DELETE_CHILD Then
        Wscript.Echo vbTab & "Delete child objects"
    End If

    If objACE.AccessMask And ADS_RIGHT_ACTRL_DS_LIST Then
        Wscript.Echo vbTab & "List child objects"
    End If

    If objACE.AccessMask And ADS_RIGHT_DS_SELF Then
        Wscript.Echo vbTab & "Perform an operation controlled by a validated write access right"
    End If

    If objACE.AccessMask And ADS_RIGHT_DS_READ_PROP Then
        Wscript.Echo vbTab & "Read properties"
    End If

    If objACE.AccessMask And ADS_RIGHT_DS_WRITE_PROP Then
        Wscript.Echo vbTab & "Write properties"
    End If

    If objACE.AccessMask And ADS_RIGHT_DS_DELETE_TREE Then
        Wscript.Echo vbTab & "Delete all child objects"
    End If

    If objACE.AccessMask And ADS_RIGHT_DS_LIST_OBJECT Then
        Wscript.Echo vbTab & "List the object"
    End If

    If objACE.AccessMask And ADS_RIGHT_DS_CONTROL_ACCESS Then
        Wscript.Echo vbTab & "Perform an operation controlled by an extended access right"
    End If

    Wscript.Echo

Next
