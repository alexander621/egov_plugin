<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandisecommonfunctions.asp
' AUTHOR: Steve Loar
' CREATED: 04/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a collection of shared functions for merchandise. Try to keep in alphabetical order.
'
' MODIFICATION HISTORY
' 1.0   04/28/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Function CalcShippingAndHandling( iSessionID )
'------------------------------------------------------------------------------
Function CalcShippingAndHandling( iSessionID )
	Dim dAmount

'	dAmount = GetShippingAndHandlingAmount( iSessionID )
'
'	If CDbl(dAmount) > CDbl(0.00) Then
'		CalcShippingAndHandling = GetShippingAndHandlingFee( dAmount )
'	Else
'		CalcShippingAndHandling = 0.00
'	End If 

	CalcShippingAndHandling = GetShippingAndHandlingAmount( iSessionID )

End Function 


'------------------------------------------------------------------------------
' Function CalcSalesTax( iSessionID )
'------------------------------------------------------------------------------
Function CalcSalesTax( iSessionID )
	Dim dAmount, sShipToState, dSalesTax, dTaxRate, sOrgState, bInStateOnly

	dTaxRate = GetSalesTaxRate( bInStateOnly )
	If CDbl(dTaxRate) > CDbl(0.00) Then 
		If bInStateOnly Then
			sOrgState = GetOrgValue( "orgstate" )
			dAmount = GetStateSalesTaxableAmount( iSessionID, sOrgState )
		Else
			dAmount = GetSalesTaxableAmount( iSessionID )
		End If 
		If CDbl(dAmount) > CDbl(0.00) Then 
			CalcSalesTax = FormatNumber((dTaxRate * dAmount),2,,,0)
		Else
			CalcSalesTax = 0.00
		End If 
	Else
		CalcSalesTax = 0.00
	End If 

End Function 


'------------------------------------------------------------------------------
' Function CartHasItems()
'------------------------------------------------------------------------------
Function CartHasItems()
	Dim sSql, oCart

	sSql = "Select count(cartid) as hits From egov_class_cart Where sessionid = " & Session.SessionID

	Set oCart = Server.CreateObject("ADODB.Recordset")
	oCart.Open sSQL, Application("DSN"), 0, 1

	If Not oCart.EOF Then 
		If clng(oCart("hits")) > clng(0) Then
			CartHasItems = True 
		Else
			CartHasItems = False 
		End If 
	Else
		CartHasItems = False 
	End If 

	oCart.Close
	Set oCart = Nothing 

End Function 


'------------------------------------------------------------------------------------------------------------
' Sub DrawDateChoices( sName )
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""16"">Today</option>"
	response.write vbcrlf & "<option value=""17"">Yesterday</option>"
	response.write vbcrlf & "<option value=""18"">Tomorrow</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""14"">Next Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""13"">Next Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""15"">Next Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 


'-------------------------------------------------------------------------------------------------
' Function GetCartValue( iCartId, sField )
'-------------------------------------------------------------------------------------------------
Function GetCartValue( ByVal iCartId, ByVal sField )
	Dim sSql, oRs

	sSql = "SELECT " & sField & " AS selectedfield FROM egov_class_cart WHERE cartid = " & iCartId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCartValue = oRs("selectedfield")
	Else
		GetCartValue = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' Function GetItemTypeId( sItemType )
'------------------------------------------------------------------------------
Function GetItemTypeId( sItemType )
	Dim sSql, oItem

	sSql = "SELECT itemtypeid FROM egov_item_types WHERE itemtype = '" & sItemType & "'"

	Set oItem = Server.CreateObject("ADODB.Recordset")
	oItem.Open sSql, Application("DSN"), 0, 1
	
	If Not oItem.EOF Then 
		GetItemTypeId = CLng(oItem("itemtypeid"))
	Else
		GetItemTypeId = 0
	End If 
	
	oItem.close 
	Set oItem = Nothing
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetMerchandiseColor( iMerchandiseColorId )
'-------------------------------------------------------------------------------------------------
Function GetMerchandiseColor( iMerchandiseColorId )
	Dim sSql, oRs

	sSql = "SELECT merchandisecolor FROM egov_merchandisecolors "
	sSql = sSql & " WHERE merchandisecolorid = " & iMerchandiseColorId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMerchandiseColor = oRs("merchandisecolor")
	Else
		GetMerchandiseColor = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'-------------------------------------------------------------------------------------------------
' Function GetMerchandiseSize( iMerchandiseSizeId )
'-------------------------------------------------------------------------------------------------
Function GetMerchandiseSize( iMerchandiseSizeId )
	Dim sSql, oRs

	sSql = "SELECT merchandisesize FROM egov_merchandisesizes "
	sSql = sSql & " WHERE merchandisesizeid = " & iMerchandiseSizeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMerchandiseSize = oRs("merchandisesize")
	Else
		GetMerchandiseSize = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function


'-------------------------------------------------------------------------------------------------
' Function GetMerchandiseSizeDisplayOrder( iMerchandiseSizeId )
'-------------------------------------------------------------------------------------------------
Function GetMerchandiseSizeDisplayOrder( iMerchandiseSizeId )
	Dim sSql, oRs

	sSql = "SELECT displayorder FROM egov_merchandisesizes "
	sSql = sSql & " WHERE merchandisesizeid = " & iMerchandiseSizeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMerchandiseSizeDisplayOrder = oRs("displayorder")
	Else
		GetMerchandiseSizeDisplayOrder = "0"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function


'-------------------------------------------------------------------------------------------------
' Function GetSalesTaxableAmount( iSessionID )
'-------------------------------------------------------------------------------------------------
Function GetSalesTaxableAmount( iSessionID )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(C.amount),0.00) AS amount FROM egov_class_cart C, egov_item_types I "
	sSql = sSql & " WHERE C.itemtypeid = I.itemtypeid AND I.istaxable = 1 AND sessionid = " & iSessionID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetSalesTaxableAmount = CDbl(oRs("amount"))
	Else
		GetSalesTaxableAmount = 0.00
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetSalesTaxRate( bInStateOnly )
'-------------------------------------------------------------------------------------------------
Function GetSalesTaxRate( ByRef bInStateOnly )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(salestaxrate,0.00) AS salestaxrate, instateonly FROM egov_salestaxrates "
	sSql = sSql & " WHERE orgid = " & Session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetSalesTaxRate = CDbl(oRs("salestaxrate"))
		If oRs("instateonly") Then
			bInStateOnly = True 
		Else
			bInStateOnly = False 
		End If 
	Else
		GetSalesTaxRate = 0.00
		bInStateOnly = False 
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetShippingAndHandlingAmount( iSessionID )
'-------------------------------------------------------------------------------------------------
Function GetShippingAndHandlingAmount( iSessionID )
	Dim sSql, oRs, dAmount

	dAmount = CDbl(0.00)

	'sSql = "SELECT ISNULL(SUM(C.amount),0.00) AS amount FROM egov_class_cart C, egov_item_types I "
	sSql = "SELECT ISNULL(C.amount,0.00) AS amount FROM egov_class_cart C, egov_item_types I "
	sSql = sSql & " WHERE C.itemtypeid = I.itemtypeid AND I.hasshippingfees = 1 And sessionid = " & iSessionID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
'		GetShippingAndHandlingAmount = FormatNumber(CDbl(oRs("amount")),2,,,0)
		Do While Not oRs.EOF 
			' For each order get the shipping and handling for it and add together
			dAmount = dAmount + GetShippingAndHandlingFee( CDbl(oRs("amount")) )
			oRs.MoveNext 
		Loop 
	Else
		dAmount = 0.00
	End If 

	oRs.CLose
	Set oRs = Nothing 

	GetShippingAndHandlingAmount = dAmount

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetShippingAndHandlingFee( dAmount )
'-------------------------------------------------------------------------------------------------
Function GetShippingAndHandlingFee( dAmount )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(shippingfee,0.00) AS shippingfee FROM egov_merchandiseshippingfees "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " AND "
	sSql = sSql & dAmount & " BETWEEN startprice AND endprice"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetShippingAndHandlingFee = FormatNumber(CDbl(oRs("shippingfee")),2,,,0)
	Else
		GetShippingAndHandlingFee = 0.00
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' Function GetStateSalesTaxableAmount( iSessionID, sOrgState )
'-------------------------------------------------------------------------------------------------
Function GetStateSalesTaxableAmount( iSessionID, sOrgState )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(SUM(C.amount),0.00) AS amount FROM egov_class_cart C, egov_item_types I "
	sSql = sSql & " WHERE C.itemtypeid = I.itemtypeid AND I.istaxable = 1 AND sessionid = " & iSessionID
	sSql = sSql & " AND UPPER(shiptostate) = '" & UCase(sOrgState) & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetStateSalesTaxableAmount = CDbl(oRs("amount"))
	Else
		GetStateSalesTaxableAmount = 0.00
	End If 

	oRs.CLose
	Set oRs = Nothing 

End Function 



%>
