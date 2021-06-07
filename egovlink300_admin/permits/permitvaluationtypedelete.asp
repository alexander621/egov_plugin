<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitvaluationtypedelete.asp
' AUTHOR: Steve Loar
' CREATED: 04/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit valuation types
'
' MODIFICATION HISTORY
' 1.0   04/14/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitValuationTypeid, sSql

iPermitValuationTypeid = CLng(request("permitvaluationtypeid") )

If PermitValuationTypeExists( iPermitValuationTypeid ) Then 
	' Clear out any step table entries
	sSql = "DELETE FROM egov_permitvaluationtypestepfees WHERE permitvaluationtypeid = " & iPermitValuationTypeid
	RunSQL sSql
	' Clear out anything in the fee types to fixture types table
	sSql = "DELETE FROM egov_permitfeetypes_to_permitvaluationtypes WHERE permitvaluationtypeid = " & iPermitValuationTypeid
	RunSQL sSql
	' Clear out the fixture entry
	sSql = "DELETE FROM egov_permitvaluationtypes WHERE permitvaluationtypeid = " & iPermitValuationTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql
End If 

response.redirect "permitvaluationtypelist.asp"

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function PermitValuationTypeExists( iPermitFixtureTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitValuationTypeExists( iPermitFixtureTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitvaluationtypeid) AS hits FROM egov_permitvaluationtypes WHERE permitvaluationtypeid = " & iPermitValuationTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitValuationTypeExists = True 
		Else
			PermitValuationTypeExists = False 
		End If 
	Else
		PermitValuationTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
