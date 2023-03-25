Set objArgs = Wscript.Arguments
If objArgs.Count <> 5 AND objArgs.Count <> 4 AND objArgs.Count <> 3 Then
   Help()
   WSCript.Quit 1
End If

Dim func : func = Lcase(objArgs(0))
Dim FileName : FileName = objArgs(1)
Dim arg1 : arg1 = objArgs(2)
Dim arg2 : arg2 = objArgs(3)
if objArgs.Count > 4 then Dim arg3 : arg3 = objArgs(4)
TristateTrue = 0
Errlevel = 0

Select Case func
Case "write"
   WriteIniValue FileName, arg1, arg2, arg3
Case "writedoublequotes"
   WriteIniValue FileName, arg1, arg2, """"&arg3&""""
Case "read"
   WScript.Echo ReadIniValue(FileName, arg1, arg2, "")
Case "append"
   WriteIniSection FileName, arg1, arg2, arg3
Case "insert"
   InsertString FileName, arg1, arg2
Case Else
   Help()
   Errlevel = 1
End Select
WSCript.Quit Errlevel

Function WriteIniValue(fileName, Section, KeyName, Value)
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    Dim fso, objFile, strText, strSection, strAfter, PosSection, PosEndSection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objFile = fso.OpenTextFile(fileName, ForReading, True, TristateUseDefault)
    If objFile.AtEndOfStream Then
        strText = ""
    Else
        strText = objFile.ReadAll
    End If
    objFile.Close
    set objFile = Nothing  
    WScript.Echo strText
    'Find section
    PosSection = InStr(1, strText, "[" & Section & "]", vbTextCompare)
    If PosSection>0 Then
		'Section exists. Find end of section
		PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(strText)+1
		do while Mid(strText,PosEndSection-2,2) = vbCrLf
			PosEndSection=PosEndSection-2
		Loop

		'Separate section contents
		Dim OldSection, NewSection, Line
		Dim sKeyName
		Dim Found : Found = False
		OldSection = Mid(strText, PosSection, PosEndSection - PosSection)
		OldSection = split(OldSection, vbCrLf)

		'Temp variable To find a Key
		sKeyName = LCase(KeyName & "=")

		'Enumerate section lines
		For Each Line In OldSection
			If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
				Line = KeyName & "=" & Value
				Found = True
			End If
			NewSection = NewSection & Line & vbCrLf
		Next

		If Found = False Then
			'key Not found - add it at the end of section
			
			NewSection = NewSection & KeyName & "=" & Value
		Else
			'remove last vbCrLf - the vbCrLf is at PosEndSection
			NewSection = Left(NewSection, Len(NewSection) - 2)
		End If

		'Combine pre-section, new section And post-section data.
		strText = Left(strText, PosSection-1) & NewSection & Mid(strText, PosEndSection)
    Else
        'Section Not found. Add section data at the end of file contents.
        If Right(strText, 2) <> vbCrLf And Len(strText)>0 Then 
            strText = strText & vbCrLf 
        End If
        strText = strText & "[" & Section & "]" & vbCrLf & _
        KeyName & "=" & Value
    End if
	Set objFile = fso.OpenTextFile(fileName, ForWriting, True, TristateUseDefault)
	
	objFile.Write strText
	objFile.Close
	set objFile = Nothing
End Function

Function WriteIniSection(fileName,section,content,append)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
   Dim fso, objFile, strText, strSection, strAfter, PosSection, PosEndSection
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set objFile = fso.OpenTextFile(fileName, ForReading, True, TristateUseDefault)
   If objFile.AtEndOfStream Then
        strText = ""
   Else
        strText = objFile.ReadAll
   End If
   objFile.Close
   set objFile = Nothing
   
   PosSection = InStr(1, strText, "[" & section & "]", vbTextCompare)
   If PosSection>0 Then
      'Section exists. Find end of section
      PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
      '?Is this last section?
      If PosEndSection = 0 Then PosEndSection = Len(strText)+1
	  do while Mid(strText,PosEndSection-2,2) = vbCrLf
	     PosEndSection=PosEndSection-2
	  Loop
	  strSection = Mid(strText, PosSection, PosEndSection - PosSection)
	  If StrComp(append, "true", vbTextCompare) = 0 Then
         If Right(strSection,2) <> vbCrLf Then strSection = strSection & vbCrLf
	     strSection = strSection & content
      Else
	     strSection = Left(strSection, Len(section)+4) & content & Right(strSection,Len(strSection)-Len(section)-2)
      End If
      strAfter = Left(strText, PosSection-1) & strSection & Right(strText,Len(strText)-(PosEndSection-1))
	  Set objFile = fso.OpenTextFile(fileName, ForWriting)
	  objFile.Write strAfter
	  objFile.Close
	  set objFile = Nothing
   Else
      strSection = "[" & section & "]" & vbCrLf & content
	  If Right(strText,2)=vbCrLf or strText="" Then 
	     strAfter = strText & strSection
	  Else
		 strAfter = strText & vbCrLf & strSection
	  End If
	  Set objFile = fso.OpenTextFile(fileName, ForWriting, False, TristateUseDefault)
	  objFile.Write strAfter
	  objFile.Close
	  set objFile = Nothing
   End If

   WriteIniSection = 0
End Function


Function InsertString(fileName,string1,string2)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
   Dim fso, objFile, strText, strAfter, PosSection
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set objFile = fso.OpenTextFile(fileName, ForReading, True, TristateUseDefault)
   If objFile.AtEndOfStream Then
        strText = ""
   Else
        strText = objFile.ReadAll
   End If
   objFile.Close
   set objFile = Nothing
   
   PosSection = InStr(1, strText, string1 & vbCrLf, vbTextCompare)
   If PosSection>0 Then
      'string exists. Find next row
      PosSection = PosSection + Len(string1 & vbCrLf)
      strAfter = Left(strText, PosSection-1) & string2 & vbCrLf & Right(strText,Len(strText)-(PosSection-1))
	  Set objFile = fso.OpenTextFile(fileName, ForWriting)
	  objFile.Write strAfter
	  objFile.Close
	  set objFile = Nothing
   Else
	  If Right(strText,2)=vbCrLf or strText="" Then 
	     strAfter = strText & string2
	  Else
		 strAfter = strText & vbCrLf & string2
	  End If
	  Set objFile = fso.OpenTextFile(fileName, ForWriting, False, TristateUseDefault)
	  objFile.Write strAfter
	  objFile.Close
	  set objFile = Nothing
   End If

   InsertString = 0
End Function


Function ReadIniValue(inifile,section,key,default)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
   Dim fso, objFile, strText, strSection, strValue, PosSection, PosEndSection, PosValue, PosEndValue
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set objFile = fso.OpenTextFile(inifile, ForReading, False, TristateUseDefault)
   strText = objFile.ReadAll
   objFile.Close
   set objFile = Nothing
   
   strValue = default
   'Find section
   PosSection = InStr(1, strText, "[" & section & "]", vbBinaryCompare)
   If PosSection>0 Then
      'Section exists. Find end of section
      PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
      '?Is this last section?
      If PosEndSection = 0 Then PosEndSection = Len(strText)+1
      'Separate section contents
      strSection = Mid(strText, PosSection, PosEndSection - PosSection)
	  strSection = split(strSection, vbCrLf)
	  key = key & "="
	  For Each Line In strSection
         If StrComp(Left(Line, Len(key)), key, vbTextCompare) = 0 Then
            strValue = Mid(Line, Len(key)+1)
         End If
      Next
   End If
   ReadIniValue = strValue
End Function


'Separates one field between sStart And sEnd
Function SeparateField(ByVal sFrom, ByVal sStart, ByVal sEnd)
  Dim PosB: PosB = InStr(1, sFrom, sStart, 1)
  If PosB > 0 Then
    PosB = PosB + Len(sStart)
    Dim PosE: PosE = InStr(PosB, sFrom, sEnd, 1)
    If PosE = 0 Then PosE = InStr(PosB, sFrom, vbCrLf, 1)
    If PosE = 0 Then PosE = Len(sFrom) + 1
    SeparateField = Mid(sFrom, PosB, PosE - PosB)
  End If
End Function


Sub Help()
  WScript.Echo "Syntax:" & chr(13) & chr(13) & "RWUINI.EXE [Read/Write/insert] [.INI File Name] [Section Name] [Key Name] [Volue]"
  WSCript.Quit(1)
End Sub
