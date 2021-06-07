<%
'response.write request("iCategoryID")
'response.end
Call subDeleteDiscount(request("iPriceDiscountId"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subDeleteDiscount(iPriceDiscountId)
' AUTHOR: TERRY FOSTER
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subDeleteDiscount(iPriceDiscountId)
	
	sSQL = "DELETE FROM egov_price_Discount WHERE PriceDiscountid = " &  iPriceDiscountId & ""
	'response.write sSQL
	'response.end

	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		' remove the discount
		.CommandText = sSQL
		.Execute
		' Clear out the class to discount table
		.CommandText = "DELETE FROM egov_class_to_pricediscount WHERE PriceDiscountid = " &  iPriceDiscountId & ""
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO discount page
	response.redirect "discount_mgmt.asp"

End Sub


%>
