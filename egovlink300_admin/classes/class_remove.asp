<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_remove.asp
' AUTHOR: Steve Loar
' CREATED: 03/27/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module removes items from the shopping cart and decrements the class/event size.
'
' MODIFICATION HISTORY
' 1.0   03/27/06   Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim bIsDropIn

	If request("iIsDropIn") = 1 Then
		bIsDropIn = True 
	Else
		bIsDropIn = False 
	End If 

	RemoveItemFromCart request("iCartId"), request("iTimeId"), request("sBuyOrWait"), bIsDropIn

	If OrgHasFeature("discounts") then
		' Recalculate any discounts
		DetermineDiscounts
	End If 

	response.redirect "class_cart.asp"

%>

<!-- #include file="class_global_functions.asp" //-->

<!-- #include file="../includes/common.asp" //-->


