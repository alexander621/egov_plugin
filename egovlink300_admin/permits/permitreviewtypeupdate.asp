<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitreviewtypeupdate.asp
' AUTHOR: Steve Loar
' CREATED: 01/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit review types
'
' MODIFICATION HISTORY
' 1.0   01/15/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitReviewTypeid, sSql, isBuildingPermitFee, sSuccessMsg

iPermitReviewTypeid = CLng(request("permitreviewtypeid") )

'If request("isbuildingpermittype") = "on" Then
'	isBuildingPermitFee = 1
'Else
'	isBuildingPermitFee = 0
'End If 

If iPermitReviewTypeid = CLng(0) Then 
	sSql = "INSERT INTO egov_permitreviewtypes ( orgid, reviewdescription, permitreviewtype ) "
	sSql = sSql & " VALUES ( " & session("orgid") & ", '" & dbsafe(request("reviewdescription")) 
	sSql = sSql & "', '" & dbsafe(request("permitreviewtype")) & "' )"

	iPermitReviewTypeid = RunIdentityInsert( sSql ) 

	sSuccessMsg = "Permit Review Type Created"
Else 
	sSql = "UPDATE egov_permitreviewtypes SET reviewdescription = '" & dbsafe(request("reviewdescription"))
	sSql = sSql & "', permitreviewtype = '" & dbsafe(request("permitreviewtype")) & "' "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND permitreviewtypeid = " & iPermitReviewTypeid

	RunSQL sSql 

	sSuccessMsg = "Changes Saved"
End If 

response.redirect "permitreviewtypeedit.asp?permitreviewtypeid=" & iPermitReviewTypeid & "&success=" & sSuccessMsg



%>