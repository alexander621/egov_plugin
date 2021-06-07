<!-- #include file="../includes/common.asp" //-->
<!--#Include file="merchandisecommonfunctions.asp"-->
<!-- #include file="../classes/class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandiseremove.asp
' AUTHOR: Steve Loar
' CREATED: 04/30/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module removes Merchandise purchases from the shopping cart.
'
' MODIFICATION HISTORY
' 1.0   04/30/2009   Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
response.End 
Dim iCartId, sSql, dShippingTotal, iShippingItemTypeId, iShippingCartId, iTaxItemTypeId, dSalesTax, iTaxCartId

iCartId = CLng(request("cartid"))
iUserId = GetCartValue( iCartId, "userid" )

' Remove team members
sSql = "DELETE FROM egov_class_cart_merchandiseitems WHERE cartid = " & iCartId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Remove Price 
sSql = "DELETE FROM egov_class_cart_price WHERE cartid = " & iCartId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Remove cart row
sSql = "DELETE FROM egov_class_cart WHERE cartid = " & iCartId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' You are going to need shipping and handling here - Seperate cart row - calc for merchandise only - Add for each order
iShippingItemTypeId = GetItemTypeId( "shipping and handling fees" )
' Remove any old Shipping and handling fees
sSql = "DELETE FROM egov_class_cart WHERE sessionid = " & Session.SessionID & " AND itemtypeid = " & iShippingItemTypeId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Recalc the shipping and handling on all merchandise in the cart
dShippingTotal = CalcShippingAndHandling( Session.SessionID, iOrgId )
If CDbl(dShippingTotal) > CDbl(0.00) Then
	sSql = "INSERT INTO egov_class_cart ( classid, userid, quantity, buyorwait, sessionid, "
	sSql = sSql & " orgid, isparent, attendeeuserid, isregatta, itemtypeid, isshippingfee, amount ) VALUES ( "
	sSql = sSql & "0, " & iUserId & ", 0, 'B', " & Session.SessionID
	sSql = sSql & ", " & iOrgId & ", 0, " & iUserId & ", 0," & iShippingItemTypeId & ", 1, " & dShippingTotal & " )"
	response.write sSql & "<br /><br />"
	iShippingCartId = RunIdentityInsertStatement( sSql )	
End If 


' You are going to need sales tax here - Seperate cart row - calc for merchandise only - Add for each order
iTaxItemTypeId = GetItemTypeId( "sales tax" )
' Remove any old sales tax
sSql = "DELETE FROM egov_class_cart WHERE sessionid = " & Session.SessionID & " AND itemtypeid = " & iTaxItemTypeId
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Recalc Sales Tax Here
dSalesTax = CalcSalesTax( Session.SessionID, iOrgId )
If CDbl(dSalesTax) > CDbl(0.00) Then
	sSql = "INSERT INTO egov_class_cart ( classid, userid, quantity, buyorwait, sessionid, "
	sSql = sSql & " orgid, isparent, attendeeuserid, isregatta, itemtypeid, issalestax, amount ) VALUES ( "
	sSql = sSql & "0, " & iUserId & ", 0, 'B', " & Session.SessionID
	sSql = sSql & ", " & iOrgId & ", 0, " & iUserId & ", 0," & iTaxItemTypeId & ", 1, " & dSalesTax & " )"
	response.write sSql & "<br /><br />"
	iTaxCartId = RunIdentityInsertStatement( sSql )	
End If 


' Return to cart
response.redirect "../classes/class_cart.asp"


%>

<!--#Include file="../include_top_functions.asp"-->

