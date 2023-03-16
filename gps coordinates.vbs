
' GPS Coordinates (Latitude, Longitude) from Address

address="11 rue de Turbigo Paris"

Set o = CreateObject("MSXML2.XMLHTTP")
o.open "GET", "https://maps.googleapis.com/maps/api/geocode/xml?address=" & address, False
o.send

Set objDoc = CreateObject("MSXML2.DOMDocument")
objDoc.loadXML o.responseText

WScript.Echo objDoc.selectSingleNode("GeocodeResponse/status").text
Set latlong = objDoc.selectSingleNode("GeocodeResponse/result/geometry/location")
WScript.Echo latlong.selectSingleNode("lat").text
WScript.Echo latlong.selectSingleNode("lng").text
