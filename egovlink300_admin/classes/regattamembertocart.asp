<!--#Include file="../includes/common.asp"-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: regattamembertocart.asp
' AUTHOR: Steve Loar
' CREATED: 03/11/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds new regatta members for teams to the shopping cart, 
'				or saves changes to them once in the cart.
'
' MODIFICATION HISTORY
' 1.0 03/11/2009 Steve Loar  - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iClassId, iCartId, iUserId, sSql, iMaxTeamMembers, iPriceTypeId, dUnitPrice
Dim dAmount, x, iTeamId, iItemTypeId, iTeamMemberCount

iClassId = CLng(request("classid"))
iCartId = CLng(request("cartid"))
iUserId = CLng(request("egovuserid"))
iItemTypeId = CLng(request("itemtypeid"))
iTeamId = CLng(request("teamid"))

iMaxTeamMembers = CLng(request("maxteammembers"))
iTeamMemberCount = CLng(0)
For x = 1 To iMaxTeamMembers
	If request("regattateammember" & x) <> "" Then
		iTeamMemberCount = iTeamMemberCount + CLng(1)
	End If 
Next 

iPriceTypeId = request("pricetypeid")
dUnitPrice = request("unitprice")
dAmount = FormatNumber( (CDbl(dUnitPrice) * iTeamMemberCount), 2,,,0)

If CLng(iCartId) = CLng(0) Then
	' Add a new cart item
	sSql = "INSERT INTO egov_class_cart ( classid, userid, quantity, buyorwait, sessionid, "
	sSql = sSql & " orgid, isparent, attendeeuserid, isregatta, itemtypeid, regattateamid, amount ) VALUES ( "
	sSql = sSql & iClassId & ", " & iUserId & ", " & iTeamMemberCount & ", 'B', " & Session.SessionID
	sSql = sSql & ", " & session("orgid") & ", 0, " & iUserId & ", 1," & iItemTypeId & ", " & iTeamId & ", " & dAmount & " )"
'	response.write sSql & "<br /><br />"
	iCartId = RunInsertStatement( sSql )

	' Create the pricing
	sSql = "INSERT INTO egov_class_cart_price ( cartid, pricetypeid, unitprice, amount ) VALUES ( "
	sSql = sSql & iCartId & ", " & iPriceTypeId & ", " & dUnitPrice & ", " & dAmount & " )"
'	response.write sSql & "<br /><br />"
	RunSQLStatement sSql
Else

	' Update existing cart item
	sSql = "UPDATE egov_class_cart SET "
	sSql = sSql & " userid = " & iUserId
	sSql = sSql & ", quantity = " & iTeamMemberCount
	sSql = sSql & ", attendeeuserid = " & iUserId
	sSql = sSql & ", regattateamid = " & iTeamId
	sSql = sSql & " WHERE cartid = " & iCartId
'	response.write sSql & "<br /><br />"
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

' Delete and add members
sSql = "DELETE FROM egov_class_cart_regattateammembers WHERE cartid = " & iCartId
RunSQLStatement sSql

For x = 1 To iMaxTeamMembers
	If request("regattateammember" & x) <> "" Then 
		sSql = "INSERT INTO egov_class_cart_regattateammembers ( cartid, cartteamid, regattateammember) VALUES ( "
		sSql = sSql & iCartId & ", " & iTeamId & ", '" & dbsafe(request("regattateammember" & x)) & "' )"
'		response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	End If 
Next 


' Take them to the cart
response.redirect "class_cart.asp?iuserid=" & iUserId & "&iclassid=" & iClassId & "&isregattateam=1"

%>