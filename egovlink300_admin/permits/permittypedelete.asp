<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permittypedelete.asp.asp
' AUTHOR: Steve Loar
' CREATED: 01/21/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit types
'
' MODIFICATION HISTORY
' 1.0   01/21/2008   Steve Loar - INITIAL VERSION
' 1.1	10/27/2010	Steve Loar - Changes to allow any type of permits
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitTypeid, sSql

iPermitTypeid = CLng(request("permittypeid") )

If PermitTypeExists( iPermitTypeid ) Then 

	' Clear out the permit types to fees types entry
	sSql = "DELETE FROM egov_permittypes_to_permitfeetypes WHERE permittypeid = " & iPermitTypeid 
	RunSQL sSql

	' Clear out the permit types to inspection types entry
	sSql = "DELETE FROM egov_permittypes_to_permitinspectiontypes WHERE permittypeid = " & iPermitTypeid 
	RunSQL sSql

	' Clear out the permit types to review types entry
	sSql = "DELETE FROM egov_permittypes_to_permitreviewtypes WHERE permittypeid = " & iPermitTypeid 
	RunSQL sSql

	' Clear out the permit types to alert types entry
	sSql = "DELETE FROM egov_permittypes_to_permitalerttypes WHERE permittypeid = " & iPermitTypeid 
	RunSQL sSql

	' Clear out the permit types to custom fields entry
	sSql = "DELETE FROM egov_permittypes_to_permitcustomfieldtypes WHERE permittypeid = " & iPermitTypeid 
	RunSQL sSql

	' Clear out the permit type entry
	sSql = "DELETE FROM egov_permittypes WHERE permittypeid = " & iPermitTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql

End If 

response.redirect "permittypelist.asp"

'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' boolean PermitTypeExists( iPermitTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitTypeExists( ByVal iPermitTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permittypeid) AS hits FROM egov_permittypes "
	sSql = sSql & " WHERE permittypeid = " & iPermitTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitTypeExists = True 
		Else
			PermitTypeExists = False 
		End If 
	Else
		PermitTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>
