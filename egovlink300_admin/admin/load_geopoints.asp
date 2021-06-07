<%
	Dim sSql, oPoints, counter

	response.write "Started at " & Now() & "<br />"

	counter = 0

	' Get a set of classids and timeids to be updated
	sSql = "select residentaddressid, latitude, longitude from temp_addresses"

	Set oPoints = Server.CreateObject("ADODB.Recordset")
	oPoints.Open sSQL, Application("DSN"), 0, 1

	Do While Not oPoints.EOF
		counter = counter + 1
		response.write "ID: " & oPoints("residentaddressid") & ", Lat: " & oPoints("latitude") & ", Lng: " & oPoints("longitude") & "<br />"
		If (counter Mod 20) = 0 Then 
			response.flush 
		End If 
		SetPoint oPoints("residentaddressid"), oPoints("latitude"), oPoints("longitude")
		oPoints.MoveNext
	Loop

	oPoints.close
	Set oPoints = Nothing 

	response.write "Finished at " & Now()


Sub SetPoint( iresidentaddressid, slatitude, slongitude )
	Dim sSql, oCmd 

	sSQL = "UPDATE egov_residentaddresses SET latitude = " & slatitude & ", longitude = " & slongitude
	sSQL = sSQL & " WHERE residentaddressid = " & iresidentaddressid

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing
End Sub 
%>