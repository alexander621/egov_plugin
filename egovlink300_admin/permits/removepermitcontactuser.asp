<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: removepermitcontactuser.asp
' AUTHOR: Steve Loar
' CREATED: 02/06/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Removes users from contractors. Called via AJAX from permitcontacttypeedit.asp
'
' MODIFICATION HISTORY
' 1.0   02/06/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitContactTypeId, sSql, iUserId

iPermitContactTypeId = CLng(request("permitcontacttypeid"))
iUserId = CLng(request("userid"))

If iPermitContactTypeId > CLng(0) Then 
	' Remove from the permit contact user table
	sSql = "DELETE FROM egov_permitcontacttypes_to_users WHERE permitcontacttypeid = " & iPermitContactTypeId
	sSql = sSql & " AND userid = " & iUserId
	RunSQL sSql
End If 

response.write "Success"

%>