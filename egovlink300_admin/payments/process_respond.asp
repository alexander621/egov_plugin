<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->


<%
' LOOP THRU PAYMENTS TO CHANGE STATUS FOR EACH PAYMENT CHECKED
For p = 1 to request.form("process_total")
	if request.form("process_"&p) <> "" then
		Update_Action(request.form("process_"&p))
	end if
Next


' REDIRECT USER TO ACTION LINE LIST PAGE
response.redirect "action_line_list.asp?" & request.querystring



'----------------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'----------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------
' FUNCTION UPDATE_ACTION(IID)
'----------------------------------------------------------------------------------------------------------------------
Function Update_Action(iID)

	' UPDATE STATUS FOR REQUEST
	sSQL = "UPDATE egov_payments SET paymentstatus='PROCESSED' where paymentid=" & iID
	Set oUpdate = Server.CreateObject("ADODB.Recordset")
	oUpdate.Open sSQL, Application("DSN") , 3, 1
	Set oUpdate = Nothing

	' ADD COMMENTS ROW
	AddCommentTaskComment intComment,request("external_comment"),"PROCESSED", iID,Session("UserID"),Session("OrgID")
	
End Function


'----------------------------------------------------------------------------------------------------------------------
' FUNCTION ADDCOMMENTTASKCOMMENT(SINTERNALMSG,SEXTERNALMSG,SSTATUS,IFORMID,IUSERID,IORGID)
'----------------------------------------------------------------------------------------------------------------------
Function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID)
		
		sSQL = "INSERT egov_payment_responses (action_status,action_internalcomment,action_externalcomment,action_userid,action_orgid,action_autoid) VALUES ('" & sStatus & "','" & sInternalMsg & "','" & sExternalMsg & "','" & iUserID & "','" & iOrgID & "','" &iFormID & "')"
		Set oComment = Server.CreateObject("ADODB.Recordset")
		oComment.Open sSQL, Application("DSN") , 3, 1
		Set oComment = Nothing

End Function
%>