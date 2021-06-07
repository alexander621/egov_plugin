<%@Language="VBScript"%>
<% option explicit %>
<%
' file: GeoCoder.asp
' last change Mon. Mar 13, 2005 by GCW

Const SoapServer	= "trial.serviceobjects.com"
Const SoapPath		= "/gcr/GeoCoder.asmx"
Const SoapAction	= "GetBestMatch"
Const SoapNamespace = "http://www.serviceobjects.com/"

main()

sub main()
'{
	displayPage()
'}
end sub

sub displayPage()
'{
	%>
	<html>
	<head>
		<title>ServiceObjects' DOTS GeoCoder Example for VBScript ASP 3.0</title>
	</head>
	<body>
	<form name="TheForm" id="TheForm" action="GeoCoder.asp" method="get" width="98%">

		<% displayFormBody() %>

	</form>
	<p>
	<p />	
	</body> <!-- close body and html missing in original -->
	</html>
	<%
'}
end sub

sub displayFormBody()
'{
	dim strAddress, strCity, strState, strZip, strKey
	
	strAddress = Request("Address")
	if(strAddress = "") then strAddress = "133 E De La Guerra"

	strCity = Request("City")
	if(strCity = "") then strCity = "Santa Barbara"

	strState = Request("State")
	if(strState = "") then strState = "CA"
	
	strZip = Request("Zip")
	if(strZip = "") then strZip = "93101"
	
	strKey = Request("Key")
	if(strKey = "") then strKey = ""
	
	%>
	<table cellpadding="0" cellspacing="0" border="0">
		<tr>
			<td><b>Enter an address to locate:</b></td>
		</tr>
		<tr>
			<td>Address:</td>
			<td><input type="text" name="Address" value="<%=strAddress%>"></td>
		</tr>
		<tr>
			<td>City:</td>
			<td><input type="text" name="City" value="<%=strCity%>"></td>
		</tr>
		<tr>
			<td>State:</td>
			<td><input type="text" name="State" value="<%=strState%>"></td>
		</tr>
		<tr>
			<td>Zip:</td>
			<td><input type="text" name="Zip" value="<%=strZip%>"></td>
		</tr>
		<tr>
			<td>License Key: </td>
			<td><input type="text" name="Key" value="<%=strKey%>"></td>
		</tr>
		<tr>
			<td><input type="submit" value="Submit" name="Action"><hr /></td>
		</tr>
	</table>
	
	<table>
		<tr>
			<td colspan="2">
				<% 'removed Request() use string directly
				if((strAddress <> "") and (strCity <> "") and (strState <> "") and (strZip <> "") and (strKey <> "")) then 
					sendRequest strAddress, strCity, strState, strZip, strKey 
				end if
				%>
			</td>
		</tr>
	</table>
	<%
'}
End Sub

Sub sendRequest(strAddress, strCity, strState, strZip, strKey)
'{
	dim SoapBody
	if((strAddress <> "") and (strCity <> "") and (strState <> "") and (strZip <> "") and (strKey <> "")) then 
		SoapBody = xmlSoap(strAddress, strCity, strState, strZip, strKey)
	end if
	%>
	<table border="0" cellspacing="1" cellpadding="3" rules="rows">
		<%
		if SoapBody = "" then
			%><tr><td><b>Empty Soap Body Response</b></td></tr><%
		else
			dim xml
			on error resume next
			set xml = Server.CreateObject("Microsoft.XMLDOM")
			if(err.number = 0) then		
			
				xml.async = False
				xml.loadxml(SoapBody) 
					  
				dim oNode : set oNode = xml.selectSingleNode("soap:Envelope/soap:Body/" & SoapAction & "Response/" & SoapAction & "Result")
			
				if (oNode.selectSingleNode("Error").nodeTypedValue = "") then
						%>
						<tr>
							<th colspan="2" align="left"><b>Location:</b></th>
						</tr>
						<% if(TypeName(oNode.selectSingleNode("Latitude")) = "IXMLDOMElement") then %>
						<tr>
							<td align="left">Latitude:</td>
							<td><%=oNode.selectSingleNode("Latitude").nodeTypedValue%></td>
						</tr>
						<% end if %>
						<% if(TypeName(oNode.selectSingleNode("Longitude")) = "IXMLDOMElement") then %>
						<tr>
							<td align="left">Longitude:</td>
							<td><%=oNode.selectSingleNode("Longitude").nodeTypedValue%></td>
						</tr>
						<% end if %>
  						<% if(TypeName(oNode.selectSingleNode("Zip")) = "IXMLDOMElement") then %>
						<tr>
							<td align="left">Zip + 4:</td>
							<td><%=oNode.selectSingleNode("Zip").nodeTypedValue%></td>
						</tr>
						<% end if 
						'These letters are returned in the "Level" node of the XML
						'They represent different levels of Geo Code matches. 
						'S - Street Level Matches         
						'P - Zip+4 Level Matches      
						'T - Zip+2 Level Matches      
						'Z - Zip Level Matches        
						'C - City/State Level Matches
						if(TypeName(oNode.selectSingleNode("LevelDescription")) = "IXMLDOMElement") then %>
						<tr>
							<td align="left">Level Description:</td>
							<td><%=oNode.selectSingleNode("LevelDescription").nodeTypedValue%></td>
						</tr>
						<% end if %>
				<%					
				else ' error element
				%>
					<tr>
						<td align="left" colspan="2"><b>Error</b></td>
					</tr>
					<tr>
						<td align="left">Description:</td>
						<td><%=oNode.selectSingleNode("Error").nodeTypedValue%></td>
					</tr>
				<%
				end if 'error element does not exist
			
				set xml = nothing
				
			else
				%>
					<tr>
						<td align="left">This objects requires Microsofts XML Parser 3.0 SP2 or greater.</td>
					</tr>
					<tr>
						<td>Download here: <a href="http://download.microsoft.com/download/xml/SP/3.20/W9X2KMeXP/EN-US/msxml3sp2Setup.exe">http://download.microsoft.com/download/xml/SP/3.20/W9X2KMeXP/EN-US/msxml3sp2Setup.exe</a></td>
					</tr>
				<%	
										
				Response.Write("Error: " & err.number)
	
			end if 'DOM object is valid
		
		end if 'soapbody is <> "" empty string
		%>
	</table>
	<%
'}
end sub

function xmlSoap(strAddress, strCity, strState, strZip, strKey)
'{
	' Instantiate objects to hold the XML DOM and the HTTP/XML communication:
	
	' Instantiate object for HTTP/XML communication:
	Dim xmlhttp, strSoap
	Set xmlhttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	
	' Build XML document:
	strSoap =	"<?xml version=""1.0"" encoding=""utf-8""?>" & _
				"<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
				"<soap:Body>" & _
				"<h:" & SoapAction & " xmlns:h=""" & SoapNamespace & """>" & _
				"<h:Address>" & strAddress & "</h:Address>" & _
				"<h:City>" & strCity & "</h:City>" & _
				"<h:State>" & strState & "</h:State>" & _
				"<h:PostalCode>" & strZip & "</h:PostalCode>" & _
			    "<h:LicenseKey>" & strKey & "</h:LicenseKey>" & _
				"</h:" & SoapAction & ">" & _
				"</soap:Body>" & _
				"</soap:Envelope>"

	' Build custom HTTP header:
	xmlhttp.Open "POST", "http://" & SoapServer & SoapPath, False	' False = Do not respond immediately
	xmlhttp.setRequestHeader "Man", "POST " & SoapPath & " HTTP/1.1"
	xmlhttp.setRequestHeader "Host", SoapServer
	xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
	xmlhttp.setRequestHeader "SOAPAction", SoapNamespace & SoapAction

	' Send it using the header generated above:
	xmlhttp.send(strSoap)
	'Response.Write(strSoap)

 	if xmlhttp.Status = 200 then			' Response from server was success
		xmlSoap = xmlhttp.responseText		' note xmlSoap is the function name hence the return value is a string Varient
	else									' Response from server failed
		xmlSoap = ""
		' Tell administrator what went wrong - maybe not users though
		Response.Write("Server Error...<br>")
		Response.Write("status = " & xmlhttp.status)
		Response.Write("<br>" & xmlhttp.statusText)
		Response.Write("<br><pre>" & Request.ServerVariables("ALL_HTTP") & "</pre>")
	end If

	Set xmlhttp = nothing
'}
end function
%>
