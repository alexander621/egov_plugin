<%
Call subDeleteQuestion(request("ifieldid"),request("iformid"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBGETFORMINFORMATION(IFORMID)
'--------------------------------------------------------------------------------------------------
Sub subDeleteQuestion(iFieldID,iFormID)
	
	' DELETE QUESTION
	sSQL = "DELETE FROM egov_action_form_questions WHERE questionid='" & iFieldID & "'"
	Set oDeleteQues = Server.CreateObject("ADODB.Recordset")
	oDeleteQues.Open sSQL, Application("DSN") , 3, 1
	Set oDeleteQues = Nothing

	' REDIRECT TO MANANGE FORM PAGE
	response.redirect("manage_form.asp?iformid=" & iFormID)

End Sub
%>