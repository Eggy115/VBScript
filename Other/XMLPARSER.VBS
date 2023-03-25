Dim objArgs,fso,oShell
Set objArgs = WScript.Arguments
If objArgs.Count <> 3 AND objArgs.Count <> 4 AND objArgs.Count <> 5 Then
  Wscript.Echo "Incorrect argument numbers: " & objArgs.Count
  Help()
  WSCript.Quit 1
End If

Dim FuncStr,XmlFile,XPathStr,ValueStr,OutFile
Set fso = CreateObject("Scripting.FileSystemObject")
'On Error Resume Next
'Parse command line parameter
for each Arg in Wscript.Arguments
	Select Case LCase(Trim(Mid(Arg, 1, InStr(Arg,"="))))
	Case "/func="
		FuncStr = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		FuncStr = Replace(FuncStr,"'",chr(34))
	Case "-func="
		FuncStr = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		FuncStr = Replace(FuncStr,"'",chr(34))
	Case "/xml="
		XmlFile = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		if fso.FileExists(XmlFile) <> True then
			Wscript.Echo "ERROR => " & "Could not found the " & XmlFile & " file."
			WScript.Quit(1)
		end if
	Case "-xml="
		XmlFile = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		if fso.FileExists(XmlFile) <> True then
			Wscript.Echo "ERROR => " & "Could not found the " & XmlFile & " file."
			WScript.Quit(1)
		end if		
	Case "/xpath="
		XPathStr = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		XPathStr = Replace(XPathStr,"'",chr(34))
	Case "-xpath="
		XPathStr = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		XPathStr = Replace(XPathStr,"'",chr(34))
	Case "/value="
		ValueStr = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		ValueStr = Replace(ValueStr,"'",chr(34))
	Case "-value="
		ValueStr = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
		ValueStr = Replace(ValueStr,"'",chr(34))
	Case "/output="
		OutFile = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
	Case "-output="
		OutFile = Trim(Right(Arg, Len(Arg) - InStr(Arg,"=")))
	End Select
Next
If IsEmpty(FuncStr) Then 
	Wscript.Echo "No /func defined"
	Help()
	WSCript.Quit 1
End If
If IsEmpty(XmlFile) Then 
	Wscript.Echo "No /xml defined"
	Help()
	WSCript.Quit 1
End If
If IsEmpty(XPathStr) Then 
	Wscript.Echo "No /xpath defined"
	Help()
	WSCript.Quit 1
End If



If IsEmpty(OutFile) Then OutFile = XmlFile


Dim oXmlDoc : Set oXmlDoc = CreateObject( "Msxml2.DOMDocument.6.0" )
oXmlDoc.setProperty "SelectionLanguage", "XPath"
oXmlDoc.async = False
oXmlDoc.Load XmlFile

' Return the node for XPathStr
Dim oNode : Set oNode = oXmlDoc.selectSingleNode( XPathStr )
If oNode Is Nothing Then
    Wscript.Echo "[" & FormatDateTime(Time, 3) & "] " & "Node not be found"
	WSCript.Quit(1)
End If

Select Case LCase(Trim(FuncStr))
Case "writenode"
	Wscript.Echo "Replace """ & oNode.text & """ => """ & ValueStr & """"
	oNode.text = ValueStr
	Wscript.Echo "Save in " & OutFile
	oXmlDoc.save OutFile
Case "readnode"
	Wscript.Echo "Value=" & oNode.text
Case Else
	Wscript.Echo "Unknown function: " & FuncStr
	Help()
	WSCript.Quit 1
End Select

WSCript.Quit(0)

Function Help()
	WScript.Echo vbCrLf & "HP CNB Xml Parser Utility, Version 1.00,A3" & vbCrLf & "Copyright (c) 2010 Hewlett-Packard - All Rights Reserved" & vbCrLf
	WScript.Echo "Syntax: CScript.exe /nologo xmlparser.vbs /func=<writenode|readnode> /xml=<.xml path> /xpath=<XPATH syntax> /value=<string for insteaded> /output=<.xml output file>" & vbCrLf
	WScript.Echo "Ex: CScript.exe /nologo xmlparser.vbs /func=""writenode"" /xml=""C:\System.sav\wdt\oa3pkpn.xml"" /xpath=""/OA3/ProductKeyPartNumber"" /value=""555555-612"""
	WSCript.Quit(1)
End Function
