' Check whether a file is valid UTF-8

WScript.Echo CheckValidUtf8("utf8.txt")

Function CheckValidUtf8(path)

  limit=10000
  
  Set Stream = CreateObject("ADODB.Stream")
  Stream.Type = 1 ' Binary
  Stream.Open
  Stream.LoadFromFile path
  s = Stream.Read(limit)
  Stream.Close

  l = LenB(s)

  ret=True
  
  For i=1 To l
    a = AscB(MidB(s,i,1))
    r0 = a And &H80
    r1 = a And &HC0
    r2 = a And &HE0
    r3 = a And &HF0
    r4 = a And &HF8

    If n>0 Then
      If r1=&H80 Then
        n=n-1
      Else
        ret=False
        Exit For
      End If
    Else
      If r4=&HF0 Then
        n=3
      ElseIf r3=&HE0 Then
        n=2
      ElseIf r2=&HC0 Then
        n=1
      ElseIf r0=0 Then
        n=0
      Else
        ret=False
        Exit For
      End If
    End If
  Next
    
  CheckValidUtf8=ret
End Function
