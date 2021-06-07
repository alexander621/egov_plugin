<%
Call subChangeOrder(request("iRateid"), request("sResidentType"), request("iDisplayOrder"), request("iDirection"), request("iMembershipId"), request("sMembershipType"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subChangeOrder(iRateid, sResidentType, iDisplayOrder, iDirection, iMembershipId)
' AUTHOR: Steve Loar
' CREATED: 02/03/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subChangeOrder(iRateid, sResidentType, iDisplayOrder, iDirection, iMembershipId, p_membershiptype)
	Dim iNewOrder

	iNewOrder = clng(iDisplayOrder) + clng(iDirection)
	'response.write iNewOrder & "<br />"
	
	sSQL = "SELECT rateid, displayorder "
 sSQL = sSQL & " FROM egov_poolpassrates "
 sSQL = sSQL & " WHERE orgid = " & Session("OrgID")
 sSQL = sSQL & " AND residenttype = '" & sResidentType & "' "
 sSQL = sSQL & " AND membershipid = " & iMembershipId
 sSQL = sSQL & " ORDER BY displayorder"

	If iDirection = "-1" Then
  		sSQL = sSQL & " DESC"
	End If
	
	Set oRates = Server.CreateObject("ADODB.Recordset")
	oRates.CursorLocation = 3
	oRates.Open sSQL, Application("DSN"), 3, 2
	Do While Not oRates.EOF
		'response.write oRates("displayorder") & " "
		If oRates("displayorder") = clng(iDisplayOrder) Then
			oRates("displayorder") = iNewOrder
			'response.write " NewOrder set "
		Else  
			If oRates("displayorder") = iNewOrder Then
				oRates("displayorder") = iDisplayOrder
				'response.write " Display Order set "
			End If 
		End if
		'response.write oRates("displayorder") & "<br />"
		oRates.MoveNext
	Loop 
	oRates.close
	Set oRates = nothing

	' REDIRECT TO Pool Pass Rates PAGE
	response.redirect( "poolpass_rates.asp?sResidentType=" & sResidentType & "&sMembershipType=" & p_membershiptype )

End Sub
%>
