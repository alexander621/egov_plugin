<!-- #include file="../includes/common.asp" //-->
<%
  if dbready_number(request("iformid")) then
     subDeleteForm request("iformid")
  else
    	response.redirect("list_forms.asp")
  end if
'------------------------------------------------------------------------------
sub subDeleteForm(iFormID)
 	Dim sSql, oCmd

 'Delete all form questions/answers
  sSQL = "DELETE from egov_action_form_questions WHERE formid = " & iFormID

 	set oDeleteQues = Server.CreateObject("ADODB.Recordset")
 	oDeleteQues.Open sSQL, Application("DSN"), 3, 1

	'Delete the form
 	sSQL = "DELETE FROM egov_action_request_forms WHERE action_form_id = " & iFormID

 	set oDeleteForm = Server.CreateObject("ADODB.Recordset")
 	oDeleteForm.Open sSQL, Application("DSN"), 3, 1

  set oDeleteQues = nothing
  set oDeleteForm = nothing

 	response.redirect "list_forms.asp?success=SD"

end sub
%>