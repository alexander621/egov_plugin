<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitcontacttypedelete.asp
' AUTHOR: Steve Loar
' CREATED: 01/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit contact types
'
' MODIFICATION HISTORY
' 1.0   01/30/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitContactTypeid, sSql, bIsOrganization

iPermitContactTypeid = CLng(request("permitcontacttypeid") )
bIsOrganization = GetContactTypeIsOrganization( iPermitContactTypeid )

If PermitContactTypeExists( iPermitContactTypeid ) Then 
	' Clear out anything in the permit contact types table
	sSql = "DELETE FROM egov_permitcontacttypes WHERE permitcontacttypeid = " & iPermitContactTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql
End If 

If bIsOrganization Then 
	response.redirect "permitorganizationlist.asp"
Else 
	response.redirect "permitcontactlist.asp"
End If 

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Function PermitContactTypeExists( iPermitFeeTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitContactTypeExists( iPermitContactTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitcontacttypeid) AS hits FROM egov_permitcontacttypes WHERE permitcontacttypeid = " & iPermitContactTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitContactTypeExists = True 
		Else
			PermitContactTypeExists = False 
		End If 
	Else
		PermitContactTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


%>
