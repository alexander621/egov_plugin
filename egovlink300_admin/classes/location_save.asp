<!-- #include file="../includes/common.asp" //-->
<%
Dim iLocationId, sName, sAddress1, sAddress2, sCity, sState, sZip, sMsg

iLocationId = CLng(request("iLocationId"))
sName = DBsafe( request("sName") )
sAddress1 = DBsafe( request("sAddress1") )
sAddress2 = DBsafe( request("sAddress2") )
sCity = DBsafe( request("sCity") )
sState = DBsafe( request("sState") )
sZip = DBsafe( request("sZip") )

If iLocationId = "0" Then
	' Insert new records
	sSQL = "INSERT INTO egov_class_location ( OrgID, Name, Address1, Address2, City, State, Zip ) Values ( " & Session("OrgID")
	sSql = sSql & ",'" & sName & "','" & sAddress1 & "','" & sAddress2 & "','" & sCity & "','" & sState & "','" & sZip & "' )"

	'response.write sSQL
	iLocationId = RunIdentityInsertStatement( sSql )
	sMsg = "i"
Else 
	' Update existing records
	sSQL = "UPDATE egov_class_location SET Name = '" & sName & "', Address1 = '" & sAddress1 & "', Address2 = '" & sAddress2
	sSql = sSql & "', City = '" & sCity & "', State = '" & sState & "', Zip = '" & sZip & "' WHERE Locationid = " & iLocationId

	'response.write sSQL
	RunSQLStatement sSql 
	sMsg = "s"
End If
	

' REDIRECT TO location management page
response.redirect "location_edit.asp?locationid=" & iLocationId & "&msg=" & sMsg


%>
