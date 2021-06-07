<HTML>
<HEAD>
<TITLE> E-GovLink Organization Management </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link href="../global.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
function confirm_delete(iorgid,sname)
{
	input_box=confirm("Are you sure you want to delete organization (" + iorgid + ") - " + sname + " ? \nAll parameters will be lost.");

	if (input_box==true)
		{ 
			// DELETE HAS BEEN VERIFIED
			location.href='delete_organization.asp?iorgid='+ iorgid;
		}
	else
		{
			// CANCEL DELETE PROCESS
		}
}
</script>
</HEAD>

<BODY>

<div style="margin-top:20px; margin-left:20px;" >

<font class=label>Organizations</font>

<div class="orgadminbox">
	
	<font class=label><a href="manage_organization.asp">Create a New Organization</a></font>

	<% Call subListOrganizations() %>

</div>

</div>


<!--#include file="bottom_include.asp"-->


</BODY>
</HTML>



<%
'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' SUB SUBLISTORGANIZATIONS()
'------------------------------------------------------------------------------------------------------------
Sub subListOrganizations()

	sSQL = "SELECT * FROM organizations"
	Set oOrgList = Server.CreateObject("ADODB.Recordset")
	oOrgList.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oOrgList.EOF Then

		response.write "<UL class=orglist>"

		Do while NOT oOrgList.EOF 
			response.write "<li class=orglist><a href=""manage_organization.asp?iorgid=" & oOrgList("orgid") & " "">View/Edit</a> | <a href=""javascript:confirm_delete('" & oOrgList("orgid") & "','" & UCASE(oOrgList("orgname")) & "');"">Delete</a> (" & oOrgList("orgid") & ") - " & UCASE(oOrgList("orgname"))  
			oOrgList.MoveNext
		Loop

		response.write "</UL>"

	End If
	Set oOrgList = Nothing 

End Sub
%>
