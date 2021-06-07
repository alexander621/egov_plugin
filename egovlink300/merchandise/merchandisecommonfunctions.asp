<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: merchandisecommonfunctions.asp
' AUTHOR: Steve Loar
' CREATED: 05/14/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a collection of shared functions for merchandise. Try to keep in alphabetical order.
'
' MODIFICATION HISTORY
' 1.0   05/14/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Function CalcShippingAndHandling( iSessionID, iOrgId )
'------------------------------------------------------------------------------
Function CalcShippingAndHandling( iSessionID, iOrgId )
	Dim dAmount

'	dAmount = GetShippingAndHandlingAmount( iSessionID )
'
'	If CDbl(dAmount) > CDbl(0.00) Then
'		CalcShippingAndHandling = GetShippingAndHandlingFee( dAmount, iOrgid )
'	Else
'		CalcShippingAndHandling = 0.00
'	End If 

	CalcShippingAndHandling = GetShippingAndHandlingAmount( iSessionID, iOrgId )

End Function 



'------------------------------------------------------------------------------
' Function CalcSalesTax( iSessionID, iOrgId )
'------------------------------------------------------------------------------
Function CalcSalesTax( iSessionID, iOrgId )
	Dim dAmount, sShipToState, dSalesTax, dTaxRate, sOrgState, bInStateOnly

	dTaxRate = GetSalesTaxRate( bInStateOnly, iOrgId )
	If CDbl(dTaxRate) > CDbl(0.00) Then 
		If bInStateOnly Then
			sOrgState = GetOrgValue( "orgstate", iOrgId )
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
	Dim sSql, oRs

	sSql = "SELECT itemtypeid FROM egov_item_types WHERE itemtype = '" & sItemType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetItemTypeId = CLng(oRs("itemtypeid"))
	Else
		GetItemTypeId = 0
	End If 
	
	oRs.Close 
	Set oRs = Nothing
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


'------------------------------------------------------------------------------
' Function GetMerchandisePrice( iMerchandiseCatalogId )
'------------------------------------------------------------------------------
Function GetMerchandisePrice( iMerchandiseCatalogId )
	Dim sSql, oRs

	sSql = "SELECT M.price FROM egov_merchandisecatalog C, egov_merchandise M " 
	sSql = sSql & " WHERE C.merchandiseid = M.merchandiseid AND C.merchandisecatalogid = " & iMerchandiseCatalogId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetMerchandisePrice = CDbl(oRs("price"))
	Else
		GetMerchandisePrice = CDbl(0.00)
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


'----------------------------------------------------------------------------------------
' Function GetOrgValue()
'----------------------------------------------------------------------------------------
Function GetOrgValue( ByVal sOrgField, ByVal iOrgId )
	Dim sSQL, oRs

	sSQL = "SELECT " & sOrgField & " AS orgvalue FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetOrgValue = oRs("orgvalue")
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
Function GetSalesTaxRate( ByRef bInStateOnly, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(salestaxrate,0.00) AS salestaxrate, instateonly FROM egov_salestaxrates "
	sSql = sSql & " WHERE orgid = " & iOrgId

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
Function GetShippingAndHandlingAmount( iSessionID, iOrgId )
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
			dAmount = dAmount + GetShippingAndHandlingFee( CDbl(oRs("amount")), iOrgId )
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
Function GetShippingAndHandlingFee( ByVal dAmount, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(shippingfee,0.00) AS shippingfee FROM egov_merchandiseshippingfees "
	sSql = sSql & " WHERE orgid = " & iOrgId & " AND "
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
