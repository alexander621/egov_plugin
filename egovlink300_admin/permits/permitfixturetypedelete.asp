<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitfixturetypedelete.asp.asp
' AUTHOR: Steve Loar
' CREATED: 12/19/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit fixture types
'
' MODIFICATION HISTORY
' 1.0   12/19/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitFixtureTypeid, sSql

iPermitFixtureTypeid = CLng(request("permitfixturetypeid") )

If PermitFixtureTypeExists( iPermitFixtureTypeid ) Then 
	' Clear out any step table entries
	sSql = "DELETE FROM egov_permitfixturetypestepfees WHERE permitfixturetypeid = " & iPermitFixtureTypeid
	RunSQL sSql
	' Clear out anything in the fee types to fixture types table
	sSql = "DELETE FROM egov_permitfeetypes_to_permitfixturetypes WHERE permitfixturetypeid = " & iPermitFixtureTypeid
	RunSQL sSql
	' Clear out the fixture entry
	sSql = "DELETE FROM egov_permitfixturetypes WHERE permitfixturetypeid = " & iPermitFixtureTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql
End If 

response.redirect "permitfixturetypelist.asp"

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function PermitFixtureTypeExists( iPermitFixtureTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitFixtureTypeExists( iPermitFixtureTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitfixturetypeid) AS hits FROM egov_permitfixturetypes WHERE permitfixturetypeid = " & iPermitFixtureTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitFixtureTypeExists = True 
		Else
			PermitFixtureTypeExists = False 
		End If 
	Else
		PermitFixtureTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
