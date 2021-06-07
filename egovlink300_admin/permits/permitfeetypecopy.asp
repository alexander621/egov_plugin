<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfeetypecopy.asp
' AUTHOR: Steve Loar
' CREATED: 09/10/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This copies permit formula fee types
'
' MODIFICATION HISTORY
' 1.0   09/10/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeTypeid, sSql, oRs, iIsfixturetypefee, iAtleastqty, iNotmorethanqty, iBaseamount
Dim iUnitqty, iUnitamount, iMinimumamount, iIsupfrontfee, iIsreinspectionfee, iIsbuildingpermitfee
Dim iAccountid, iIsvaluationtypefee, iIsconstructiontypefee, iOnSewerFeeReport, iIsResidentialUnitTypeFee

iPermitFeeTypeid = CLng(request("permitfeetypeid"))
sRedirectPage = request("redirectpage")

sSql = "SELECT isfixturetypefee, isvaluationtypefee, isconstructiontypefee, isresidentialunittypefee, "
sSql = sSql & " ISNULL(permitfeeprefix, '') AS permitfeeprefix, ISNULL(permitfee, '') AS permitfee, "
sSql = sSql & " permitfeecategorytypeid, permitfeemethodid, atleastqty, notmorethanqty, "
sSql = sSql & " ISNULL(baseamount,0.00) AS baseamount, unitqty, unitamount, minimumamount, ispercentagetypefee, percentage, "
sSql = sSql & " isupfrontfee, isreinspectionfee, isbuildingpermitfee, accountid, onsewerfeereport, ISNULL(upfrontamount,0.00) AS upfrontamount "
sSql = sSql & " FROM egov_permitfeetypes F "
sSql = sSql & " WHERE permitfeetypeid = " & iPermitFeeTypeid 

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	If oRs("isfixturetypefee") Then
		iIsfixturetypefee = 1
	Else
		iIsfixturetypefee = 0
	End If 
	If oRs("isvaluationtypefee") Then
		iIsvaluationtypefee = 1
	Else
		iIsvaluationtypefee = 0
	End If 
	If oRs("isconstructiontypefee") Then
		iIsconstructiontypefee = 1
	Else
		iIsconstructiontypefee = 0
	End If 
	If oRs("isresidentialunittypefee") Then
		iIsResidentialUnitTypeFee = 1
	Else
		iIsResidentialUnitTypeFee = 0
	End If 
	If IsNull(oRs("atleastqty")) Then
		iAtleastqty = "NULL"
	Else
		iAtleastqty = oRs("atleastqty")
	End If 
	If IsNull(oRs("notmorethanqty")) Then
		iNotmorethanqty = "NULL"
	Else
		iNotmorethanqty = oRs("notmorethanqty")
	End If 
	If IsNull(oRs("baseamount")) Then
		iBaseamount = "NULL"
	Else
		iBaseamount = oRs("baseamount")
	End If 
	If IsNull(oRs("unitqty")) Then
		iUnitqty = "NULL"
	Else
		iUnitqty = oRs("unitqty")
	End If 
	If IsNull(oRs("unitamount")) Then
		iUnitamount = "NULL"
	Else
		iUnitamount = oRs("unitamount")
	End If 
	If IsNull(oRs("minimumamount")) Then
		iMinimumamount = "NULL"
	Else
		iMinimumamount = oRs("minimumamount")
	End If
	If oRs("isupfrontfee") Then
		iIsupfrontfee = 1
	Else
		iIsupfrontfee = 0
	End If 
	If oRs("isreinspectionfee") Then
		iIsreinspectionfee = 1
	Else
		iIsreinspectionfee = 0
	End If 
	If oRs("isbuildingpermitfee") Then
		iIsbuildingpermitfee = 1
	Else
		iIsbuildingpermitfee = 0
	End If 
	If IsNull(oRs("accountid")) Then
		iAccountid = "NULL"
	Else
		iAccountid = oRs("accountid")
	End If
	If oRs("ispercentagetypefee") Then
		iIspercentagetypefee = 1
	Else
		iIspercentagetypefee = 0
	End If 

	If IsNull(oRs("percentage")) Then
		sPercentage =  "NULL"
	Else 
		sPercentage = CDbl(oRs("percentage"))
	End If 

	If oRs("onsewerfeereport") Then
		iOnSewerFeeReport = 1
	Else
		iOnSewerFeeReport = 0
	End If

	'orgid, isfixturetypefee, isvaluationtypefee, isconstructiontypefee, ispercentagetypefee, permitfeeprefix, permitfee, 
    'permitfeecategorytypeid, permitfeemethodid, atleastqty, notmorethanqty, baseamount, unitqty, unitamount, percentage, 
	'minimumamount, isupfrontfee, isreinspectionfee, isbuildingpermitfee, accountid

	sSql = "INSERT INTO egov_permitfeetypes ( orgid, isfixturetypefee, isvaluationtypefee, isconstructiontypefee, ispercentagetypefee, "
	sSql = sSql & " permitfeeprefix, permitfee, permitfeecategorytypeid, permitfeemethodid, atleastqty, notmorethanqty, baseamount, "
	sSql = sSql & " unitqty, unitamount, percentage, minimumamount, isupfrontfee, isreinspectionfee, isbuildingpermitfee, accountid, "
	sSql = sSql & " onsewerfeereport, upfrontamount, isresidentialunittypefee ) VALUES ( "
	sSql = sSql & session("orgid") & ", " & iIsfixturetypefee & ", " & iIsvaluationtypefee & ", " & iIsconstructiontypefee & ", " & iIspercentagetypefee & ", '"
	sSql = sSql & oRs("permitfeeprefix") & "', 'Copy of " & oRs("permitfee") & "', " & oRs("permitfeecategorytypeid") & ", "
	sSql = sSql & oRs("permitfeemethodid") & ", " & iAtleastqty & ", " & iNotmorethanqty & ", "& iBaseamount & ", " 
	sSql = sSql & iUnitqty & ", " & iUnitamount & ", " & sPercentage & ", " & iMinimumamount & ", " & iIsupfrontfee & ", "
	sSql = sSql & iIsreinspectionfee & ", " & iIsbuildingpermitfee & ", " & iAccountid & ", " & iOnSewerFeeReport & ", "
	sSql = sSql & oRs("upfrontamount") & ", " & iIsResidentialUnitTypeFee & " )" 
'	response.write sSql & "<br />"

	iNewPermitFeeTypeId = RunIdentityInsert( sSql )

	If oRs("isfixturetypefee") Then
		' Get the fixtures and bring them over
		CopyPermitFixtures iPermitFeeTypeid, iNewPermitFeeTypeId
	End If 

	If oRs("isvaluationtypefee") Then
		' Copy the valuation step fees
		CopyPermitValuationStepFees iPermitFeeTypeid, iNewPermitFeeTypeId
	End If 

	If oRs("isresidentialunittypefee") Then
		' Copy the Residential Unit step fees
		CopyPermitResidentialUnitStepFees iPermitFeeTypeid, iNewPermitFeeTypeId
	End If 

	' Pull in any Fee Multipliers
	CopyPermitFeeMultipliers iPermitFeeTypeid, iNewPermitFeeTypeId

End If 

oRs.Close
Set oRs = Nothing 

response.redirect sRedirectPage & ".asp?permitfeetypeid=" & iNewPermitFeeTypeId & "&success=Copy%20Succeeded"


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Sub CopyPermitFixtures( iPermitFeeTypeId, iNewPermitFeeTypeId )
'-------------------------------------------------------------------------------------------------
Sub CopyPermitFixtures( iPermitFeeTypeId, iNewPermitFeeTypeId )
	Dim sSql, oRs, iPermitFixtureTypeId

	' Get the fixtures for the old fee type, if any exist
	sSql = "SELECT F.permitfixturetypeid, F.permitfixture, T.displayorder "
	sSql = sSql & " FROM egov_permitfixturetypes F, egov_permitfeetypes_to_permitfixturetypes T "
	sSql = sSql & " WHERE T.permitfixturetypeid = F.permitfixturetypeid AND T.permitfeetypeid = " & iPermitFeeTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		'  permitfixturetypeid, orgid, permitfixture
		sSql = "INSERT INTO egov_permitfixturetypes ( orgid, permitfixture ) VALUES ( "
		sSql = sSql & session("orgid") & ", '" & dbsafe(oRs("permitfixture"))& "' )"
		response.write sSql & "<br />"
		iPermitFixtureTypeId = RunIdentityInsert( sSql )

		'  permitfeetypeid, permitfixturetypeid, displayorder
		sSql = "INSERT INTO egov_permitfeetypes_to_permitfixturetypes ( permitfeetypeid, permitfixturetypeid, displayorder ) "
		sSql = sSql & " VALUES ( " & iNewPermitFeeTypeId & ", " & iPermitFixtureTypeId & ", " & oRs("displayorder") & " )"
		RunSQL sSql

		' Input the step table entries for each fixture
		CopyFixtureStepFees iPermitFixtureTypeId, oRs("permitfixturetypeid")

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'-------------------------------------------------------------------------------------------------
' Sub CopyFixtureStepFees( iPermitFixtureTypeId, iOldPermitFixtureTypeId )
'-------------------------------------------------------------------------------------------------
Sub CopyFixtureStepFees( iPermitFixtureTypeId, iOldPermitFixtureTypeId )
	Dim sSql, oRs

	sSql = "SELECT fixturetypestepfeeid, ISNULL(atleastqty, 0) AS atleastqty, ISNULL(notmorethanqty, 999999999) AS notmorethanqty, "
	sSql = sSql & " ISNULL(baseamount,0.00) AS baseamount, ISNULL(unitqty,1) AS unitqty, ISNULL(unitamount,0.00) AS unitamount "
	sSql = sSQl & " FROM egov_permitfixturetypestepfees WHERE permitfixturetypeid = " & iOldPermitFixtureTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		'  fixturetypestepfeeid, permitfixturetypeid, atleastqty, notmorethanqty, baseamount, unitqty, unitamount
		sSql = "INSERT INTO egov_permitfixturetypestepfees ( permitfixturetypeid, atleastqty, notmorethanqty, "
		sSql = sSql & " baseamount, unitqty, unitamount ) VALUES ( " & iPermitFixtureTypeId & ", "
		sSql = sSql & oRs("atleastqty") & ", " & oRs("notmorethanqty") & ", " & oRs("baseamount") & ", "
		sSql = sSql & oRs("unitqty") & ", " & oRs("unitamount") & " )"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'-------------------------------------------------------------------------------------------------
' Sub CopyPermitFeeMultipliers( iPermitFeeTypeId, iNewPermitFeeTypeId )
'-------------------------------------------------------------------------------------------------
Sub CopyPermitFeeMultipliers( iPermitFeeTypeId, iNewPermitFeeTypeId )
	Dim sSql, oRs, iPermitMultiplierTypeId

	sSql = "SELECT F.feemultipliertypeid, F.feemultiplier, F.feemultiplierrate, T.displayorder "
	sSql = sSql & " FROM egov_feemultipliertypes F, egov_permitfeetypes_to_feemultipliertypes T "
	sSql = sSql & " WHERE F.feemultipliertypeid = T.feemultipliertypeid AND T.permitfeetypeid = " & iPermitFeeTypeId
	sSql = sSql & " ORDER BY T.displayorder, F.feemultipliertypeid"

	'response.write "<p> Multiplier: " & sSql & "</p><br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		'  feemultipliertypeid, orgid, feemultiplier, feemultiplierrate
		sSql = "INSERT INTO egov_feemultipliertypes ( orgid, feemultiplier, feemultiplierrate ) VALUES ( " 
		sSql = sSql & session("orgid") & ", '" & dbsafe(oRs("feemultiplier")) & "', " & oRs("feemultiplierrate")
		sSql = sSql & " )"
		response.write sSql & "<br />"
		iPermitMultiplierTypeId = RunIdentityInsert( sSql )

		' permitfeetypeid, feemultipliertypeid, displayorder
		sSql = "INSERT INTO egov_permitfeetypes_to_feemultipliertypes ( permitfeetypeid, feemultipliertypeid, displayorder ) "
		sSql = sSql & " VALUES ( " & iNewPermitFeeTypeId & ", " & iPermitMultiplierTypeId & ", " & oRs("displayorder") & " )"
		RunSQL sSql

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' Sub CopyPermitValuationStepFees( iPermitFeeTypeId, iNewPermitFeeTypeId )
'-------------------------------------------------------------------------------------------------
Sub CopyPermitValuationStepFees( iPermitFeeTypeId, iNewPermitFeeTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(atleastvalue, 0.00) AS atleastvalue, ISNULL(notmorethanvalue, 999999999.99) AS notmorethanvalue, "
	sSql = sSql & " ISNULL(baseamount,0.00) AS baseamount, ISNULL(unitqty,1) AS unitqty, ISNULL(unitamount,0.00) AS unitamount "
	sSql = sSQl & " FROM egov_permitvaluationtypestepfees WHERE permitfeetypeid = " & iPermitFeeTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		' valuationtypestepfeeid, permitvaluationtypeid, atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount
		sSql = "INSERT INTO egov_permitvaluationtypestepfees ( permitfeetypeid, "
		sSql = sSql & " atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount ) VALUES ( "
		sSql = sSql & iNewPermitFeeTypeId & ", " & oRs("atleastvalue") & ", " & oRs("notmorethanvalue") & ", "
		sSql = sSql & oRs("baseamount") & ", " & oRs("unitqty") & ", " & oRs("unitamount") & " )"
		'response.write sSql & "<br />"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'-------------------------------------------------------------------------------------------------
' Sub CopyPermitResidentialUnitStepFees( iPermitFeeTypeid, iNewPermitFeeTypeId )
'-------------------------------------------------------------------------------------------------
Sub CopyPermitResidentialUnitStepFees( iPermitFeeTypeid, iNewPermitFeeTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(atleastqty, 0) AS atleastqty, ISNULL(notmorethanqty, 999999999) AS notmorethanqty, "
	sSql = sSql & " ISNULL(baseamount,0.00) AS baseamount, ISNULL(unitqty,1) AS unitqty, ISNULL(unitamount,0.00) AS unitamount "
	sSql = sSQl & " FROM egov_permitresidentialunittypestepfees WHERE permitfeetypeid = " & iPermitFeeTypeid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		' valuationtypestepfeeid, permitvaluationtypeid, atleastvalue, notmorethanvalue, baseamount, unitqty, unitamount
		sSql = "INSERT INTO egov_permitresidentialunittypestepfees ( permitfeetypeid, "
		sSql = sSql & " atleastqty, notmorethanqty, baseamount, unitqty, unitamount ) VALUES ( "
		sSql = sSql & iNewPermitFeeTypeId & ", " & oRs("atleastqty") & ", " & oRs("notmorethanqty") & ", "
		sSql = sSql & oRs("baseamount") & ", " & oRs("unitqty") & ", " & oRs("unitamount") & " )"
		'response.write sSql & "<br />"
		RunSQL sSql
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub 



%>
