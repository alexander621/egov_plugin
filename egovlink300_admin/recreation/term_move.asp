<%
Call subChangeOrder(request("iFacilityId"), request("iDisplayOrder"), request("iDirection"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subChangeOrder(iFacilityID, iDisplayOrder, iDirection)
' AUTHOR: Steve Loar
' CREATED: 01/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subChangeOrder(iFacilityID, iDisplayOrder, iDirection)
	Dim iNewOrder

	iNewOrder = clng(iDisplayOrder) + clng(iDirection)
	//response.write iNewOrder & "<br />"
	
	sSQL = "Select termid, displayorder FROM egov_recreation_terms where facilityid = " & iFacilityId & " order by displayorder"
	If iDirection = "-1" Then
		sSql = sSql & " DESC"
	End If
	
	Set oTerms = Server.CreateObject("ADODB.Recordset")
	oTerms.CursorLocation = 3
	oTerms.Open sSQL, Application("DSN"), 3, 2
	Do While Not oTerms.EOF
		//response.write oterms("displayorder") & " "
		If oterms("displayorder") = clng(iDisplayOrder) Then
			oterms("displayorder") = iNewOrder
			//response.write " NewOrder set "
		Else  
			If oterms("displayorder") = iNewOrder Then
				oterms("displayorder") = iDisplayOrder
				//response.write " Display Order set "
			End If 
		End if
		//response.write oterms("displayorder") & "<br />"
		oTerms.MoveNext
	Loop 
	oTerms.close
	Set oTerms = nothing

	' REDIRECT TO Facility terms PAGE
	response.redirect( "facility_terms.asp?facilityid=" & iFacilityId )

End Sub
%>
