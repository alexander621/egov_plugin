<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>

	<title>Hanahan, South Carolina</title>

	<link rel="stylesheet" type="text/css" href="http://hanahan.besavvy.egovlink.com/primarytemplatefiles/custom.css" />

</head>
<body>

	<p style="font-size:11px;" align="center">
<%

	'Format Time

	Dim oXMLDoc, hnow

	response.Write(MonthName(Month(Now()), 1) & " " & Day(Now()) & ", " & Year(Now()))

	If Hour(Now()) > 12 Then 
		hnow = Hour(Now()) - 12
	Else
		hnow = Hour(Now())
	End If 

	sMinute = Minute(Now())
	If sMinute < 10 Then 
		sMinute = "0" & sMinute
	End If 

	response.Write(" &middot; " & hnow & ":" & sMinute & "<br />")

	'Get Weather

	Set oXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")
	oXMLDoc.setProperty "ServerHTTPRequest", true
	oXMLDoc.async=False

	'oXMLDoc.load "http://weather.gov/data/current_obs/KCHS.xml"
	oXMLDoc.load "http://www.weather.gov/xml/current_obs/KCHS.xml"

	If oXMLDoc.parseError.reason <> "" Then 
		Response.Write "XML load failed: " & oXMLDoc.parseError.reason
	Else 
		'Response.Write "XML load succeeded!"
		response.Write("Currently in Hanahan: <br />")
		response.Write(oXMLDoc.getElementsByTagName("weather").item(0).text)
		response.Write(", ")
		response.Write(oXMLDoc.getElementsByTagName("temp_f").item(0).text)
		response.write("&deg;")
	End If 

%>
	</p>

</body>
</html>