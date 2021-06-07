<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<!-- #include file="../includes/JSON_2.0.2.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: feepickeradd.asp
' AUTHOR: Steve Loar
' CREATED: 04/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Adds fees to permits
'
' MODIFICATION HISTORY
' 1.0   04/30/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iPermitFeeTypeId, sSql, oRs, iIsfixturetypefee, iAtleastqty, iNotmorethanqty, iBaseamount, iUnitqty, iUnitamount
Dim iMinimumamount, iIsupfrontfee, iIsreinspectionfee, iIsbuildingpermitfee, iIsrequired, iDisplayOrder, iOnSewerFeeReport
Dim iPermitFeeId, iAccountid, iFeeAmount, iIsvaluationtypefee, iIsconstructiontypefee, iIspercentagetypefee, sPercentage
Dim bOnBBSFeeReport, sResponse, iIsResidentialUnitTypeFee, iFeeReportingTypeId

' Create the JSON object to pass data back to the calling page
Set sResponse = jsObject()

iPermitId = CLng(request("permitid"))

iPermitFeeTypeId = CLng(request("permitfeetypeid"))

iDisplayOrder = GetNextDisplayOrder( iPermitId )
iFeeAmount = CDbl(0.00) 

sSql = "SELECT F.isfixturetypefee, F.isvaluationtypefee, F.isconstructiontypefee, "
sSql = sSql & " ISNULL(F.permitfeeprefix, '') AS permitfeeprefix, ISNULL(F.permitfee, '') AS permitfee, "
sSql = sSql & " F.permitfeecategorytypeid, F.permitfeemethodid, F.atleastqty, F.notmorethanqty, "
sSql = sSql & " ISNULL(F.baseamount,0.00) AS baseamount, F.unitqty, F.unitamount, F.minimumamount, "
sSql = sSql & " F.isupfrontfee, F.isreinspectionfee, F.isbuildingpermitfee, F.accountid, M.isflatfee, "
sSql = sSql & " F.ispercentagetypefee, F.percentage, ISNULL(F.upfrontamount,0.00) AS upfrontamount, "
sSql = sSql & " ISNULL(F.feereportingtypeid,0) AS feereportingtypeid, F.isresidentialunittypefee "
sSql = sSql & " FROM egov_permitfeetypes F, egov_permitfeemethods M "
sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.permitfeetypeid = " & iPermitFeeTypeId 

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
	iIsrequired = 0

	If oRs("isflatfee") Then
		iFeeAmount = CDbl(oRs("baseamount"))
	Else
		iFeeAmount = CDbl(0.00)
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

	If oRs("isresidentialunittypefee") Then
		iIsResidentialUnitTypeFee = 1
	Else
		iIsResidentialUnitTypeFee = 0
	End If 

	If CLng(oRs("feereportingtypeid")) <> CLng(0) Then 
		iFeeReportingTypeId = CLng(oRs("feereportingtypeid"))
	Else
		iFeeReportingTypeId = "NULL"
	End If 

	sSql = "INSERT INTO egov_permitfees ( permitid, permitfeetypeid, orgid, isfixturetypefee, isvaluationtypefee, isconstructiontypefee, permitfeeprefix, permitfee, "
	sSql = sSql & " permitfeecategorytypeid, permitfeemethodid, atleastqty, notmorethanqty, baseamount, quantity, unitqty, "
	sSql = sSql & " unitamount, minimumamount, isupfrontfee, isreinspectionfee, isbuildingpermitfee, accountid, isrequired, "
	sSql = sSql & " displayorder, feeamount, amountpaid, includefee, ispercentagetypefee, percentage, upfrontamount, feereportingtypeid, isresidentialunittypefee ) VALUES ( " & iPermitId & ", " & iPermitFeeTypeId & ", "
	sSql = sSql & session("orgid") & ", " & iIsfixturetypefee & ", " & iIsvaluationtypefee & ", " & iIsconstructiontypefee & ", '" & dbsafe(oRs("permitfeeprefix")) & "', '" & dbsafe(oRs("permitfee")) & "', "
	sSql = sSql & oRs("permitfeecategorytypeid") & ", " & oRs("permitfeemethodid") & ", " & iAtleastqty & ", " & iNotmorethanqty & ", "
	sSql = sSql & iBaseamount & ", 0, " & iUnitqty & ", " & iUnitamount & ", " & iMinimumamount & ", " & iIsupfrontfee & ", "
	sSql = sSql & iIsreinspectionfee & ", " & iIsbuildingpermitfee & ", " & iAccountid & ", " & iIsrequired & ", "
	sSql = sSql & iDisplayOrder & ", " & iFeeAmount & ", 0.00, 1, " & iIspercentagetypefee & ", " & sPercentage & ", "
	sSql = sSql & oRs("upfrontamount") & ", " & iFeeReportingTypeId & ", " & iIsResidentialUnitTypeFee & " )"
	iPermitFeeId = RunIdentityInsert( sSql )

	If oRs("isfixturetypefee") Then
		' Get the fixtures and bring them over
		'response.write "<p>Fixtures Here</p>"
		CreatePermitFixtures iPermitId, iPermitFeeTypeId, iPermitFeeId
	End If 

	If oRs("isvaluationtypefee") Then
		' Get the valuations and bring them over
		CreatePermitValuationStepFees iPermitId, iPermitFeeTypeId, iPermitFeeId
	End If 

	If oRs("isresidentialunittypefee") Then
		CreatePermitResidentialUnitStepFees iPermitId, iPermitFeeTypeId, iPermitFeeId
	End If 

	' Pull in any Fee Multipliers
	CreatePermitFeeMultipliers iPermitFeeTypeId, iPermitFeeId, iPermitId

	' Format the JSON return
	sResponse("flag") = "success"
	sResponse("permitfeeid") = iPermitFeeId
	sResponse("permitfee") = oRs("permitfee")
	sResponse("permitfeeprefix") = oRs("permitfeeprefix")
	sResponse("permitfeemethod") = GetPermitFeeMethodById( oRs("permitfeemethodid") )

Else
	' Format the JSON return
	sResponse("flag") = "failed"
End If  

oRs.Close
Set oRs = Nothing 



sResponse.Flush
'Set sResponse = Nothing 


'-------------------------------------------------------------------------------------------------
' Function GetNextDisplayOrder( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetNextDisplayOrder( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(MAX(displayorder),0) AS displayorder FROM egov_permitfees WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetNextDisplayOrder = CLng(oRs("displayorder")) + CLng(1)
	Else
		GetNextDisplayOrder = CLng(1)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 



%>
