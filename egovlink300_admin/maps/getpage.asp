<% 
	sStreet = "4303 Hamilton Ave"
	sRepStreet = Replace(sStreet," ","+")
	sCity = "Cincinnati"
	sState = "OH"
	sZip = "45223"
	checkaddress = sStreet & ","& sCity & "," & sState & "," & sZip
    url = "http://rpc.geocoder.us/service/csv?address=" & sStreet & "+" & sCity & "+" & sState & "+" & sZip  
    set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
    xmlhttp.open "GET", url, false 
    xmlhttp.send "" 
    Response.write xmlhttp.responseText & "<br />"
	sText = xmlhttp.responseText
	If InStr(sText, ",") > 0  And InStr(sText, checkaddress) > 0 Then 
		aValues=Split(sText,",",-1,1)
		response.write "Lat: " & aValues(0) & "<br />"
		response.write "Lng: " & aValues(1) & "<br />"
	Else
		Response.write sText
	End If 
    set xmlhttp = nothing 
%>