<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: RATE_SAVE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE FACILITY RATE AND ADD NEW RATES
'
' MODIFICATION HISTORY
' 1.0   01/17/06	JOHN STULLENBERGER - INITIAL VERSION
' 1.0   01/18/06	STEVE LOAR - CODE ADDED
' 2.0	01/22/07	JOHN STULLENBERGER - NEW VERSION WITH DIFFERENT PRICING LEVELS
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRateId, sSql

iRateId = CLng(request("iRateId"))

If iRateId = CLng(0) Then
	' insert the data
	sSql = "INSERT INTO egov_facility_rates (orgid, ratedescription) VALUES ( " & Session("OrgID") & ", '" & dbsafe( request("ratedescription") ) & "' )" 
	iRateID = RunInsertStatement( sSql )
Else
	' update the desctiption
	sSql = "UPDATE egov_facility_rates SET ratedescription = '" & dbsafe( request("ratedescription") ) & "' WHERE rateid = " & iRateId & " AND orgid = " & Session("OrgID")

	RunSQLStatement sSql
End If 


' ADD PRICES FOR RATE ID
subSavePrices iRateID

' REDIRECT TO FACILITY RATES PAGE
response.redirect( "facility_rates.asp" )


'------------------------------------------------------------------------------------------------------------
' SUB SUBSAVEPRICES(IRATEID)
'------------------------------------------------------------------------------------------------------------
Sub subSavePrices( ByVal iRateID )
	Dim sSql, iPriceTypeCount, x, iPriceTypeId, Amount

	' clear out any old rate types and prices
	sSql = "DELETE FROM egov_facility_rate_to_pricetype WHERE rateid = " & iRateId 
	RunSQLStatement sSql

	' now get the count of rate types 
	iPriceTypeCount = CLng(request("pricetypecount") )

	' loop through and add them
	For x = 1 To iPriceTypeCount

		iPriceTypeId = CLng(request("pricetypeid" & x))
		' amount is supposed to be currency, but the table only accepts whole numbers 
		If request("amount" & x) <> "" Then 
			Amount = CLng(request("amount" & x))
		Else
			Amount = 0.00
		End If 

		sSql = "INSERT INTO egov_facility_rate_to_pricetype ( rateid, pricetypeid, amount ) VALUES ( " & iRateId & ", " & iPriceTypeId & ", " & Amount & " )"
		RunSQLStatement sSql

	Next 

End Sub
%>