<!--#Include file="../includes/common.asp"-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: regattateamtocart.asp
' AUTHOR: Steve Loar
' CREATED: 03/04/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds regatta teams to the shopping cart, 
'				or saves changes to them once in the cart.
'
' MODIFICATION HISTORY
' 1.0 03/04/2009 Steve Loar  - Initial Version
' 1.1	04/07/2010	Steve Loar - No more regatta team members, added team group size
' 1.2	5/14/2010	Steve Loar - Split captain name into first and last
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iClassId, iCartId, iUserId, sTeamName, sCaptainFirstName, sCaptainLastName, sCaptainAddress, sCaptainCity
Dim sSql, sCaptainState, sCaptainZip, sCaptainPhone, iMaxTeamMembers, iTotalMembers
Dim iPriceTypeId, dUnitPrice, dAmount, x, iCartTeamId, iItemTypeId, iRegattaTeamGroupId

iClassId = CLng(request("classid"))
iCartId = CLng(request("cartid"))
iUserId = CLng(request("egovuserid"))
iItemTypeId = CLng(request("itemtypeid"))

sTeamName = "'" & dbsafe(request("regattateam")) & "'"
iRegattaTeamGroupId = CLng(request("regattateamgroupid"))
sCaptainFirstName = "'" & dbsafe(request("captainfirstname")) & "'"
sCaptainLastName = "'" & dbsafe(request("captainlastname")) & "'"
sCaptainAddress = "'" & dbsafe(request("captainaddress")) & "'"
sCaptainCity = "'" & dbsafe(request("captaincity")) & "'"
sCaptainState = "'" & dbsafe(request("captainstate")) & "'"
sCaptainZip = "'" & dbsafe(request("captainzip")) & "'"
sCaptainPhone = "'" & dbsafe(request("captainphone")) & "'"

'iMaxTeamMembers = CLng(request("maxteammembers"))
'iTeamMemberCount = CLng(0)
'For x = 1 To iMaxTeamMembers
'	If request("regattateammember" & x) <> "" Then
'		iTeamMemberCount = iTeamMemberCount + CLng(1)
'	End If 
'Next 
'iTotalMembers = iTeamMemberCount + CLng(1) ' Add the teammembers and the captain to get price multiplier

iTotalMembers = 1  ' Now just the captain.

iPriceTypeId = request("pricetypeid")
dUnitPrice = request("unitprice")
dAmount = FormatNumber( (CDbl(dUnitPrice) * iTotalMembers), 2,,,0)

If CLng(iCartId) = CLng(0) Then
	' Add a new cart item
	sSql = "INSERT INTO egov_class_cart ( classid, userid, quantity, buyorwait, sessionid, "
	sSql = sSql & " orgid, isparent, attendeeuserid, isregatta, itemtypeid ) VALUES ( "
	sSql = sSql & iClassId & ", " & iUserId & ", " & iTotalMembers & ", 'B', " & Session.SessionID
	sSql = sSql & ", " & session("orgid") & ", 0, " & iUserId & ", 1," & iItemTypeId & " )"
	iCartId = RunInsertStatement( sSql )

	' Create the pricing
	sSql = "INSERT INTO egov_class_cart_price ( cartid, pricetypeid, unitprice, amount ) VALUES ( "
	sSql = sSql & iCartId & ", " & iPriceTypeId & ", " & dUnitPrice & ", " & dAmount & " )"
	RunSQLStatement sSql
Else
	' Update existing cart item
	sSql = "UPDATE egov_class_cart SET "
	sSql = sSql & " userid = " & iUserId
	sSql = sSql & ", quantity = " & iTotalMembers
	sSql = sSql & ", attendeeuserid = " & iUserId
	sSql = sSql & " WHERE cartid = " & iCartId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql

	' Update the pricing
	sSql = "UPDATE egov_class_cart_price SET "
	sSql = sSql & " pricetypeid = " & iPriceTypeId
	sSql = sSql & ", unitprice = " & dUnitPrice
	sSql = sSql & ", amount = " & dAmount
	sSql = sSql & " WHERE cartid = " & iCartId
'	response.write sSql & "<br /><br />"
	RunSQLStatement sSql
End If 

' Delete and add the team information
sSql = "DELETE FROM egov_class_cart_regattateams WHERE cartid = " & iCartId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

sSql = "INSERT INTO egov_class_cart_regattateams ( cartid, regattateam, orgid, classid, captainfirstname, captainlastname, "
sSql = sSql & " captainaddress, captaincity, captainstate, captainzip, captainphone, regattateamgroupid ) VALUES ( "
sSql = sSql & iCartId & ", " & sTeamName & ", " & session("orgid") & ", " & iClassId & ", " 
sSql = sSql & sCaptainFirstName & ", " & sCaptainLastName & ", " & sCaptainAddress & ", " & sCaptainCity & ", "
sSql = sSql & sCaptainState & ", " & sCaptainZip & ", " & sCaptainPhone & ", " & iRegattaTeamGroupId & " )"
response.write sSql & "<br /><br />"
iCartTeamId = RunInsertStatement( sSql )

' Delete and add members
'sSql = "DELETE FROM egov_class_cart_regattateammembers WHERE cartid = " & iCartId
'RunSQLStatement sSql

'For x = 1 To iMaxTeamMembers
'	If request("regattateammember" & x) <> "" Then 
'		sSql = "INSERT INTO egov_class_cart_regattateammembers ( cartid, cartteamid, regattateammember) VALUES ( "
'		sSql = sSql & iCartId & ", " & iCartTeamId & ", '" & dbsafe(request("regattateammember" & x)) & "' )"
'		'response.write sSql & "<br /><br />"
'		RunSQLStatement sSql
'	End If 
'Next 


' Take them to the cart
response.redirect "class_cart.asp?iuserid=" & iUserId & "&iclassid=" & iClassId & "&isregattateam=1"


%>