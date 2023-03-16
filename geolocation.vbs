' Geolocation of IP Address ( IP to Country) in IIS Logs

d=Now
limit=d-5

Dim fso
Set fso = CreateObject("Scripting.Filesystemobject")

Dim dic
Set dic = CreateObject("Scripting.Dictionary")

Set ip = CreateObject("Vbsedit.iptocountry")

If Not(fso.FileExists("iptocountry.bin")) Then
  If fso.FileExists("iptocountry.csv") Then
    ip.CreateDatabaseFromCsv "iptocountry.csv","iptocountry.bin"
  Else
    WScript.Quit
  End If
Else
  If fso.FileExists("iptocountry.csv") Then
    If fso.GetFile("iptocountry.csv").DateLastModified > _
          fso.GetFile("iptocountry.bin").DateLastModified Then
      ip.CreateDatabaseFromCsv "iptocountry.csv","iptocountry.bin"
    End If
  End If
End If

'if you need to convert a lot of IP addresses
ip.LoadDatabaseIntoMemory "iptocountry.bin"

Dim theDate,theTime,serviceName,serverName,serverIP
Dim method,uriStem,uriQuery,serverPort,username,clientIp,protocolVersion
Dim userAgent,cookie,referrer,host,protocolStatus
Dim subStatus,win32Status,bytesSentByServer,bytesReceived,timeTaken
  
Do While d>limit

  file = "C:\LogFiles\W3SVC4\u_ex" & mydatepart("yyyy",d) & mydatepart("m",d) & myDatepart("d",d) & ".log"

  If Not(fso.FileExists(file)) then
    Exit Do
  End if

  Dim ff
  Set ff = fso.OpenTextFile(file,1)

  Do while not(ff.AtEndofstream)
    Dim str
    str = ff.ReadLine
  
    filterok=False
  
    If left(str,9)="#Fields: " Then
      ParseFieldNames split(Mid(str,10)," ")
    ElseIf Left(str,1)<>"#" Then
      filterok=True
      ParseFields split(str," ")
    End If
  
    If filterok Then
      code = ip.IpToCountry(clientIp)
      country = ip.CountryCodeToCountryName(code)
      WScript.Echo country
      WScript.Echo theDate,theTime,serviceName,serverName,serverIP
      WScript.Echo method,uriStem,uriQuery,serverPort,username,clientIp,protocolVersion
      WScript.Echo userAgent,cookie,referrer,host,protocolStatus
      WScript.Echo subStatus,win32Status,bytesSentByServer,bytesReceived,timeTaken
    End If
  Loop

  d=d-1
Loop

Sub ParseFields(arr)
  theDate=""
  If dic.Exists("date") Then
    theDate= arr(dic.Item("date"))
  End If
 
  theTime=""
  If dic.Exists("time") Then
    theTime= arr(dic.Item("time"))
  End If
  
  serviceName=""
  If dic.Exists("s-sitename") Then
    serviceName= arr(dic.Item("s-sitename"))
  End If

  serverName=""
  If dic.Exists("s-computername") Then
    serverName= arr(dic.Item("s-computername"))
  End If

  serverIP=""
  If dic.Exists("s-ip") Then
    serverIP= arr(dic.Item("s-ip"))
  End If

  method=""
  If dic.Exists("cs-method") Then
    method= arr(dic.Item("cs-method"))
  End If

  uriStem=""
  If dic.Exists("cs-uri-stem") Then
    uriStem= arr(dic.Item("cs-uri-stem"))
  End If

  uriQuery=""
  If dic.Exists("cs-uri-query") Then
    uriQuery= arr(dic.Item("cs-uri-query"))
  End If

  serverPort=""
  If dic.Exists("s-port") Then
    serverPort= arr(dic.Item("s-port"))
  End If

  username=""
  If dic.Exists("cs-username") Then
    username= arr(dic.Item("cs-username"))
  End If

  clientIp=""
  If dic.Exists("c-ip") Then
    clientIp= arr(dic.Item("c-ip"))
  End If

  protocolVersion=""
  If dic.Exists("cs-version") Then
    protocolVersion= arr(dic.Item("cs-version"))
  End If

  userAgent=""
  If dic.Exists("cs(User-Agent)") Then
    userAgent= arr(dic.Item("cs(User-Agent)"))
  End If

  cookie=""
  If dic.Exists("cs(Cookie)") Then
    cookie= arr(dic.Item("cs(Cookie)"))
  End If

  referrer=""
  If dic.Exists("cs(Referrer)") Then
    referrer= arr(dic.Item("cs(Referrer)"))
  End If

  host=""
  If dic.Exists("cs-host") Then
    host= arr(dic.Item("cs-host"))
  End If

  protocolStatus=""
  If dic.Exists("sc-status") Then
    protocolStatus= arr(dic.Item("sc-status"))
  End If

  subStatus=""
  If dic.Exists("sc-substatus") Then
    subStatus= arr(dic.Item("sc-substatus"))
  End If

  win32Status=""
  If dic.Exists("sc-win32-status") Then
    win32Status= arr(dic.Item("sc-win32-status"))
  End If

  bytesSentByServer=0
  If dic.Exists("sc-bytes") Then
    bytesSentByServer= arr(dic.Item("sc-bytes"))
  End If

  bytesReceived=0
  If dic.Exists("cs-bytes") Then
    bytesReceived= arr(dic.Item("cs-bytes"))
  End If

  timeTaken=0
  If dic.Exists("time-taken") Then
    timeTaken= arr(dic.Item("time-taken"))
  End If
 
End Sub

Sub ParseFieldNames(arr)
  dic.RemoveAll
  For i=0 to UBound(arr)
    dic.Add arr(i),i
  Next
End Sub

Function Mydatepart(attr,d)
  Dim v
  v = CStr(DatePart(attr,d))
  if Len(v)=1 then
    v = "0" & v
  elseif len(v)>2 then
     v = Right(v,2)
  end if
  Mydatepart=v
End Function
