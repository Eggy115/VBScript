option explicit
On Error GoTo 0

Dim strResult: strResult = Wscript.ScriptName
Dim strExpr, strInExp, strLastY, yyy
strInExp = "1+1"
strLastY = ""
Do While True

  strExpr = inputbox("Last calculation:" & vbCR & strLastY, "Calculator", strInExp)

  If Len( strExpr) = 0 Then Exit Do 

  ''' in my locale, decimal separator is a comma but VBScript arithmetics premises a dot  
  strExpr = Replace( strExpr, ",", ".")   ''' locale specific

  On Error Resume Next             ' enable error handling
  yyy = Eval( strExpr)
  If Err.Number = 0 Then
    strInExp = CStr( yyy)
    strLastY = strExpr & vbTab & strInExp
    strResult = strResult & vbNewLine & strLastY
  Else
    strLastY = strExpr & vbTab & "!!! 0x" & Hex(Err.Number) & " " & Err.Description
    strInExp = strExpr
    strResult = strResult & vbNewLine & strLastY
  End If
  On Error GoTo 0                  ' disable error handling
Loop

Wscript.Echo strResult
Wscript.Quit
