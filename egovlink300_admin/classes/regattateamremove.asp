<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattateamremove.asp
' AUTHOR: Steve Loar
' CREATED: 03/06/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module removes Regatta Teams from the shopping cart.
'
' MODIFICATION HISTORY
' 1.0   03/06/2009   Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iCartId, sSql

iCartId = CLng(request("cartid"))

' Remove team members
sSql = "DELETE FROM egov_class_cart_regattateammembers WHERE cartid = " & iCartId
'response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Remove team
sSql = "DELETE FROM egov_class_cart_regattateams WHERE cartid = " & iCartId
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

' Return to cart
response.redirect "class_cart.asp"


%>