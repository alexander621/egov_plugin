<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: removepermitfee.asp
' AUTHOR: Steve Loar
' CREATED: 05/01/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Removes fees from permits
'
' MODIFICATION HISTORY
' 1.0   05/01/2008	Steve Loar - INITIAL VERSION
' 1.1	07/14/2009	Steve Loar - Put check for empty parameter
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeId, sSql

If request("permitfeeid") <> "" Then 

	iPermitFeeId = CLng(request("permitfeeid"))

	' Remove from the permit fee table
	sSql = "DELETE FROM egov_permitfees WHERE permitfeeid = " & iPermitFeeId
	RunSQL sSql

	' Remove from the fixtures table
	sSql = "DELETE FROM egov_permitfixtures WHERE permitfeeid = " & iPermitFeeId
	RunSQL sSql

	' Remove from the fixtures step table
	sSql = "DELETE FROM egov_permitfixturestepfees WHERE permitfeeid = " & iPermitFeeId
	RunSQL sSql

	' Remove from the valuations table
	'sSql = "DELETE FROM egov_permitvaluations WHERE permitfeeid = " & iPermitFeeId
	'RunSQL sSql

	' Remove from the valuations step table
	sSql = "DELETE FROM egov_permitvaluationstepfees WHERE permitfeeid = " & iPermitFeeId
	RunSQL sSql

	' Remove from the residential unit step table
	sSql = "DELETE FROM egov_permitresidentialunitstepfees WHERE permitfeeid = " & iPermitFeeId
	RunSQL sSql

	' Remove from the multipliers table
	sSql = "DELETE FROM egov_permitfeemultipliers WHERE permitfeeid = " & iPermitFeeId
	RunSQL sSql

End If 

response.write "Success"

%>
