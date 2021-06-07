<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfeetypedelete.asp
' AUTHOR: Steve Loar
' CREATED: 01/09/2008
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit fee types
'
' MODIFICATION HISTORY
' 1.0   01/09/2008   Steve Loar - INITIAL VERSION
' 1.1	04/14/2008	Steve Loar - Valuation Fees Added
' 1.2	11//3/2008	Steve Loar - Residential Unit step fees added
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFeeTypeid, sSql

iPermitFeeTypeid = CLng(request("permitfeetypeid") )

If PermitFeeTypeExists( iPermitFeeTypeid ) Then 
	' Clear out anything in the fee types to fixture types table
	sSql = "DELETE FROM egov_permitfeetypes_to_permitfixturetypes WHERE permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql
	' Clear out anything in the fee types to multiplier types table
	sSql = "DELETE FROM egov_permitfeetypes_to_feemultipliertypes WHERE permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql
	' Clear out anything in the fee types to valuation types table
	sSql = "DELETE FROM egov_permitfeetypes_to_permitvaluationtypes WHERE permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql
	' Clear out anything in the residential unit step fees table
	sSql = "DELETE FROM egov_permitresidentialunittypestepfees WHERE permitfeetypeid = " & iPermitFeeTypeid
	RunSQL sSql
	' Clear out the fee types entry
	sSql = "DELETE FROM egov_permitfeetypes WHERE permitfeetypeid = " & iPermitFeeTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql
End If 

response.redirect "permitfeetypelist.asp"

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function PermitFeeTypeExists( iPermitFeeTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitFeeTypeExists( iPermitFeeTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitfeetypeid) AS hits FROM egov_permitfeetypes WHERE permitfeetypeid = " & iPermitFeeTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitFeeTypeExists = True 
		Else
			PermitFeeTypeExists = False 
		End If 
	Else
		PermitFeeTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
