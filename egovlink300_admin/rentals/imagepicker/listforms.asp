<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link href="../global.css" rel="stylesheet" type="text/css">
</HEAD>

<BODY>
<p><% Call subListForms() %></p>
</BODY>
</HTML>


<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' SUB SUBLISTFORMS()
'------------------------------------------------------------------------------------------------------------
Sub subListForms()

	sSQL = "SELECT * FROM egov_action_request_forms where orgid='" & session("orgid") & "' order by action_form_type,action_form_name"
	Set oFormList = Server.CreateObject("ADODB.Recordset")
	oFormList.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oFormList.EOF Then

		response.write "<table cellspacing=0 cellpadding=2 >"

		Do while NOT oFormList.EOF 
			
			sFormNameOnly = replace(UCASE(oFormList("action_form_name")),"'","\'")
			sFormNameOnly = replace(sFormNameOnly,chr(10),"")
			sFormNameOnly = replace(sFormNameOnly,chr(13),"")
			sFormNameOnly = Trim(sFormNameOnly)

			response.write "<tr style=""cursor:hand;"" onClick=""parent.document.frmAddArticle.AFormName.value='" & sFormNameOnly  & "';parent.document.frmAddArticle.iFormID.value=" & oFormList("action_form_id") & """ ><td > (" & oFormList("action_form_id") & ") </td><td>" & UCASE(oFormList("action_form_name")) & "</td></tr>" 
			oFormList.MoveNext
		Loop

		response.write "</table>"

	End If
	Set oFormList = Nothing 

End Sub
%>