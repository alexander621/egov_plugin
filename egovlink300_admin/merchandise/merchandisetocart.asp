<!--#Include file="../includes/common.asp"-->
<!--#Include file="merchandisecommonfunctions.asp"-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: merchandisetocart.asp
' AUTHOR: Steve Loar
' CREATED: 04/30/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds merchandise to the shopping cart, 
'				or saves changes to them once in the cart.
'
' MODIFICATION HISTORY
' 1.0 04/30/2009 Steve Loar  - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iCartId, iUserId, iItemTypeId, sShipToName, sShipToAddress, sShipToCity, sShipToState
Dim sShipToZip, iTotalItems, iPriceTypeId, dAmount, iMerchandiseItemCount, x, dShippingTotal
Dim iShippingItemTypeId, iTaxItemTypeId, iShippingCartId, dSalesTax, iTaxCartId

iCartId = CLng(request("cartid"))
iUserId = CLng(request("egovuserid"))
Session("eGovUserId") = iUserId
iItemTypeId = CLng(request("itemtypeid"))
iMerchandiseItemCount = CLng(request("maxmerchandiseitems"))
dShippingTotal = CDbl(0.00)

sShipToName = "'" & dbsafe(request("shiptoname")) & "'"
sShipToAddress = "'" & dbsafe(request("shiptoaddress")) & "'"
sShipToCity = "'" & dbsafe(request("shiptocity")) & "'"
sShipToState = "'" & dbsafe(request("shiptostate")) & "'"
sShipToZip = "'" & dbsafe(request("shiptozip")) & "'"

' Set the price and number of items
dAmount = CDbl(0.00)
iTotalItems = CLng(0)

iPriceTypeId = "0" ' There are no price types for merchandise but the table requires one


If CLng(iCartId) = CLng(0) Then
	' Add a new cart item
	sSql = "INSERT INTO egov_class_cart ( classid, userid, quantity, buyorwait, sessionid, "
	sSql = sSql & " orgid, isparent, attendeeuserid, isregatta, itemtypeid ) VALUES ( "
	sSql = sSql & "0, " & iUserId & ", 0, 'B', " & Session.SessionID
	sSql = sSql & ", " & session("orgid") & ", 0, " & iUserId & ", 0," & iItemTypeId & " )"
	'response.write sSql & "<br /><br />"
	iCartId = RunInsertStatement( sSql )

	' Create the pricing
	sSql = "INSERT INTO egov_class_cart_price ( cartid, pricetypeid, unitprice, amount ) VALUES ( "
	sSql = sSql & iCartId & ", " & iPriceTypeId & ", NULL, 0.00 )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
Else 
	' Remove any old merchandise items
	sSql = "DELETE FROM egov_class_cart_merchandiseitems WHERE cartid = " & iCartId
	RunSQLStatement sSql
End If 

'Put the merchandise items in
For x = 1 To iMerchandiseItemCount 
	If CLng(request("quantity" & x)) > CLng(0) Then 
		dAmount = dAmount + CDbl(CDbl(request("price" & x)) * CLng(request("quantity" & x)))
		iTotalItems = iTotalItems + CLng(request("quantity" & x))
		sSql = "INSERT INTO egov_class_cart_merchandiseitems ( cartid, orgid, merchandisecatalogid, quantity, price) VALUES ( "
		sSql = sSql & iCartId & ", " & session("orgid") & ", " & request("merchandisecatalogid" & x) & ", " & request("quantity" & x) & ", " & request("price" & x) & " )"
		'response.write sSql & "<br /><br />"
		RunSQLStatement sSql
	End If 
Next 


' Update existing cart item
sSql = "UPDATE egov_class_cart SET "
sSql = sSql & " userid = " & iUserId
sSql = sSql & ", quantity = " & iTotalItems
sSql = sSql & ", attendeeuserid = " & iUserId
sSql = sSql & ", amount = " & dAmount
sSql = sSql & ", shiptoname = " & sShipToName
sSql = sSql & ", shiptoaddress = " & sShipToAddress
sSql = sSql & ", shiptocity = " & sShipToCity
sSql = sSql & ", shiptostate = " & sShipToState
sSql = sSql & ", shiptozip = " & sShipToZip
sSql = sSql & " WHERE cartid = " & iCartId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Update the pricing
sSql = "UPDATE egov_class_cart_price SET "
sSql = sSql & "amount = " & dAmount
sSql = sSql & " WHERE cartid = " & iCartId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql


' You are going to need shipping and handling here - Seperate cart row - calc for merchandise only - Add for each order
iShippingItemTypeId = GetItemTypeId( "shipping and handling fees" )
' Remove any old Shipping and handling fees
sSql = "DELETE FROM egov_class_cart WHERE sessionid = " & Session.SessionID & " AND itemtypeid = " & iShippingItemTypeId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Recalc the shipping and handling on all merchandise in the cart sum into one row in cart
dShippingTotal = CalcShippingAndHandling( Session.SessionID )
If CDbl(dShippingTotal) > CDbl(0.00) Then
	sSql = "INSERT INTO egov_class_cart ( classid, userid, quantity, buyorwait, sessionid, "
	sSql = sSql & " orgid, isparent, attendeeuserid, isregatta, itemtypeid, isshippingfee, amount ) VALUES ( "
	sSql = sSql & "0, " & iUserId & ", 0, 'B', " & Session.SessionID
	sSql = sSql & ", " & session("orgid") & ", 0, " & iUserId & ", 0," & iShippingItemTypeId & ", 1, " & dShippingTotal & " )"
	'response.write sSql & "<br /><br />"
	iShippingCartId = RunInsertStatement( sSql )	
End If 


' You are going to need sales tax here - Seperate cart row - calc for merchandise only - Add for each order
iTaxItemTypeId = GetItemTypeId( "sales tax" )
' Remove any old sales tax
sSql = "DELETE FROM egov_class_cart WHERE sessionid = " & Session.SessionID & " AND itemtypeid = " & iTaxItemTypeId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' recalc the tax on all merchandise in the cart
dSalesTax = CalcSalesTax( Session.SessionID )
If CDbl(dSalesTax) > CDbl(0.00) Then
	sSql = "INSERT INTO egov_class_cart ( classid, userid, quantity, buyorwait, sessionid, "
	sSql = sSql & " orgid, isparent, attendeeuserid, isregatta, itemtypeid, issalestax, amount ) VALUES ( "
	sSql = sSql & "0, " & iUserId & ", 0, 'B', " & Session.SessionID
	sSql = sSql & ", " & session("orgid") & ", 0, " & iUserId & ", 0," & iTaxItemTypeId & ", 1, " & dSalesTax & " )"
	'response.write sSql & "<br /><br />"
	iTaxCartId = RunInsertStatement( sSql )	
End If 


' Take them to the cart
response.redirect "../classes/class_cart.asp?iuserid=" & iUserId 

%>