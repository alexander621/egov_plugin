<!-- #include file="../includes/common.asp" //-->

<%

	subSaveDiscount request("iPriceDiscountId"), request("iClass"), request("sName"), request("sQTYRequired"), request("sAmount"), request("sDescription"), request("bIsShared"), request("discounttypeid")
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subSaveDiscount(iPriceDiscountId, iClass, sName, sQTYRequired, sAmount, sDescription, bIsShared)
' AUTHOR: Terry Foster
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subSaveDiscount(iPriceDiscountId, iClass, sName, sQTYRequired, sAmount, sDescription, bIsShared, iDiscountTypeId)

	sName = DBsafe( sName )
	sDescription = DBsafe( sDescription )

	if not isnumeric(sQTYRequired) then sQTYRequired = 0
	if not isnumeric(sAmount) then sAmount = 0

	if bIsShared = "on" then 
		bIsShared = 1
	else
		bIsShared = 0
	end if


	If iPriceDiscountId = "0" Then
		' Insert new records
		sSQL = "INSERT INTO egov_price_discount (OrgID,ClassID,DiscountName,qtyrequired,DiscountAmount,DiscountDescription,isshared, discounttypeid ) Values (" & Session("OrgID") & ",'" & iClass & "','" & sName & "','" & sQTYRequired & "',CONVERT(money," & sAmount & "),'" & sDescription & "'," & bIsShared & ", " & iDiscountTypeId & ")"
	Else 
		' Update existing records
		sSQL = "UPDATE egov_price_discount SET ClassID='" & iClass & "', DiscountName='" & sName & "', qtyrequired='" & sQTYRequired & "', DiscountAmount=CONVERT(Money," & sAmount & "), DiscountDescription='" & sDescription & "', isshared=" & bIsShared & ", discounttypeid = " & iDiscountTypeId & " WHERE PriceDiscountid = " & iPriceDiscountId & ""
	End If


	'response.write sSQL
	'response.end
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSQL
		.Execute
	End With
	Set oCmd = Nothing
	

	' REDIRECT TO discount management page
	response.redirect "discount_mgmt.asp"

End Sub
%>
