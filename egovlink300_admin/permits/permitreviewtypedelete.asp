<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitreviewtypedelete.asp
' AUTHOR: Steve Loar
' CREATED: 01/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the permit Review types
'
' MODIFICATION HISTORY
' 1.0   01/15/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitReviewTypeid, sSql

iPermitReviewTypeid = CLng(request("permitreviewtypeid") )

If PermitReviewTypeExists( iPermitReviewTypeid ) Then 
	' Clear out the permit types to review types entry
	sSql = "DELETE FROM egov_permittypes_to_permitreviewtypes WHERE permitreviewtypeid = " & iPermitReviewTypeid 
	RunSQL sSql
	' Clear out the review type entry
	sSql = "DELETE FROM egov_permitreviewtypes WHERE permitreviewtypeid = " & iPermitReviewTypeid & " AND orgid = " & session("orgid")
	RunSQL sSql
End If 

response.redirect "permitreviewtypelist.asp"

'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' boolean PermitReviewTypeExists( iPermitReviewTypeid )
'-------------------------------------------------------------------------------------------------
Function PermitReviewTypeExists( ByVal iPermitReviewTypeid )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitreviewtypeid) AS hits FROM egov_permitreviewtypes "
	sSql = sSql & " WHERE permitreviewtypeid = " & iPermitReviewTypeid
	sSql = sSql & " AND orgid = "& session("orgid" )

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitReviewTypeExists = True 
		Else
			PermitReviewTypeExists = False 
		End If 
	Else
		PermitReviewTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>
