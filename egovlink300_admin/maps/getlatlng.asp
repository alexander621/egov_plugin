<% 
	sStreet = "201 E 4th St"
	sRepStreet = Replace(sStreet," ","+")
	sCity = "Cincinnati"
	sState = "OH"
	sZip = "45219"
	checkaddress = sStreet & "," & sCity & "," & sState 
    url = "http://rpc.geocoder.us/service/csv?address=" & sRepStreet & "+" & sCity & "+" & sState '& "+" & sZip  
	'response.write url & "<br />"
    set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
    xmlhttp.open "GET", url, false 
    xmlhttp.send "" 
    'Response.write xmlhttp.responseText & "<br />"
	sText = xmlhttp.responseText
	If InStr(sText, ",") > 0  And InStr(sText, sStreet) > 0 Then 
		aValues=Split(sText,",",-1,1)
		response.write "Lat: " & aValues(0) & "<br />"
		response.write "Lng: " & aValues(1) & "<br />"
	Else
		Response.write sText
	End If 
    set xmlhttp = nothing 
%>