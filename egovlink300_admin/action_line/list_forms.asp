<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "form creator" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 
%>

<html>
<head>
<title> E-GovLink Forms Management </title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="JavaScript">
<!--
	function confirm_delete(iorgid,sname)
	{
		input_box=confirm("Are you sure you want to delete form (" + iorgid + ") - " + sname + " ? \nAll parameters will be lost.");

		if (input_box==true)
			{ 
				// DELETE HAS BEEN VERIFIED
				location.href='delete_form.asp?iorgid='+ iorgid;
			}
		else
			{
				// CANCEL DELETE PROCESS
			}
	}
//-->
</script>

</head>

<body>

<div style="margin-top:20px; margin-left:20px;" >

<font class=label>Forms - List View </font>

<% 'blnCanEditForms = HasPermission("CanEditActionForms") 
	blnCanEditForms = True 
%>

<% If  blnCanEditForms Then %>
<div class="orgadminboxf">
	
	<font class=label><a href="manage_form.asp">Create a New Form</a></font>

	<p><% subListForms %></p>

</div>

<%Else%>

	<p>You do not have permission to access the <b>E-Gov Forms Creator section</b>.  Please contact your E-Govlink administrator to inquire about gaining access to the <b>E-Gov Forms Creator section</b>.</p>

<%End If%>
	 

</div>


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

	sSQL = "SELECT * FROM egov_action_request_forms order by action_form_type,action_form_name"
	Set oFormList = Server.CreateObject("ADODB.Recordset")
	oFormList.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oFormList.EOF Then

		response.write "<table cellspacing=0 cellpadding=2 class=formlist>"
		response.write "<tr class=formrowheader ><td>Actions</td><td colspan=2>Form Name</td><td>Type</td></tr>"

		Do while NOT oFormList.EOF 
			' SET ROW BG COLOR
			If sRowClass = "formrowG" Then
				sRowClass = "formrowW" 
			Else
				sRowClass = "formrowG"
			End If

			' DETERMINE FORM TYPE
			If oFormList("action_form_type") = "1" Then
				sType = "CUSTOM"
			Else
				sType = "STANDARD"
			End If

			' DETERMINE IF FORM IS AVAILABLE
			If oFormList("action_form_enabled") Then
				sEnabled = "Disable"
			Else
				sEnabled = "Enable"
			End If
			
			
			response.write "<tr class=" & sRowClass & "><td class=formlist><a href=""?iorgid=" & oFormList("action_form_id") & " "">" & sEnabled & "</a> | <a href=""manage_form.asp?iformid=" & oFormList("action_form_id") & " "">View/Edit</a> | <a href=""javascript:confirm_delete('" & oFormList("action_form_id") & "','" & UCASE(oFormList("action_form_name")) & "');"">Delete</a></td><td class=formlist> (" & oFormList("action_form_id") & ") </td><td class=formlist class=formlist>" & UCASE(oFormList("action_form_name")) & "</td><td class=formlist>" & sType & "</td></tr>" 
			oFormList.MoveNext
		Loop

		response.write "</table>"

	End If
	Set oFormList = Nothing 

End Sub
%>
