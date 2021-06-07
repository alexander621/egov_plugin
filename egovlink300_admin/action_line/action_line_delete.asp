<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: action_line_delete.asp
' AUTHOR: Steve Loar
' CREATED: 11/15/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module deletes Action Line Requests.
'
' MODIFICATION HISTORY
' 1.0	11/15/06	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim sSql, oCmd, ActionId

	ActionId = CLng(request("id"))

	If ActionId > 0 Then 
 	 	Set oCmd = Server.CreateObject("ADODB.Command")
	 	 With oCmd
   			.ActiveConnection = Application("DSN")

   		'Delete the responses
			   .CommandText = "DELETE FROM egov_action_responses WHERE action_autoid = " & ActionId 
   			.Execute

   		'Delete the issue locations
			   .CommandText = "DELETE FROM egov_action_response_issue_location WHERE actionrequestresponseid = " & ActionId
   			.Execute

   		'Delete the attachments
			   .CommandText = "DELETE FROM egov_submitted_request_attachments WHERE submitted_request_id = " & ActionId
   			.Execute

   		'Delete the code sections
			   .CommandText = "DELETE FROM egov_submitted_request_code_sections WHERE submitted_request_id = " & ActionId
   			.Execute

   		'Delete the custom questions/answers (public and internal only)
			   .CommandText = "DELETE FROM egov_submitted_request_fields WHERE submitted_request_id = " & ActionId
   			.Execute

   		'Delete the fees
			   .CommandText = "DELETE FROM egov_action_fees WHERE action_autoid = " & ActionId
   			.Execute

			  'Delete the request
   			.CommandText = "DELETE FROM egov_actionline_requests WHERE action_autoid = " & ActionId 
   			.Execute

		End With 

		Set oCmd = Nothing
	End If 

	response.redirect "action_line_list.asp"
%>