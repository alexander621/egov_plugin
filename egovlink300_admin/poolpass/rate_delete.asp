<%
Call subDeleteRate(request("sResidentType"), request("iRateid"), request("sMembershipType"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subDeleteRate(sResidentType, iRateid)
' AUTHOR: Steve Loar
' CREATED: 01/31/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subDeleteRate(sResidentType, iRateid, p_membershiptype)
	
	' Delete from the facilitytimepart table
	sSQL = "DELETE FROM egov_poolpassrates WHERE rateid = " & iRateid  & ""

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
		' delete from the preselected family member table
		.CommandText = "Delete from egov_poolpasspreselected where rateid = " & iRateid  & ""
		.Execute
	End With
	Set oCmd = Nothing

	' Reset the display order
	subReOrder sResidentType 

	' REDIRECT TO Pool Pass Rates PAGE
	response.redirect( "poolpass_rates.asp?sResidentType=" & sResidentType & "&sMembershipType=" & p_membershiptype & "&success=SD")

End Sub

Sub subReOrder( sResidentType )
	Dim iNewOrder

	iNewOrder = 0
	
	sSQL = "Select rateid, displayorder FROM egov_poolpassrates where orgid = " & Session("OrgID") & " and residenttype = '" & sResidentType & "' order by displayorder"
	
	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.CursorLocation = 3
	oRates.Open sSQL, Application("DSN"), 3, 2

	Do While Not oRates.EOF
		iNewOrder = iNewOrder + 1
		oRates("displayorder") = iNewOrder
		oRates.MoveNext
	Loop 
	oRates.close
	Set oRates = nothing

	' REDIRECT TO Pool Pass Rates PAGE
	'response.redirect( "poolpass_rates.asp?sResidentType=" & sResidentType )

End Sub
%>