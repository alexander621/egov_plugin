<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
if request.querystring("export") <> "" then
sDate = Month(Date()) & Day(Date()) & Year(Date())
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=EGOV_Do_Not_Knock_" & request.querystring("export") & "_EXPORT_" & sDate & ".xls"

	sSQL = BuildQuery(dbSafe(request.querystring("export")))
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1

	if not oRs.EOF then 
		ListMembers oRs
	end if
	oRs.Close
	Set oRs = Nothing
	response.end
end if
if request.ServerVariables("REQUEST_METHOD") = "POST" then
	sSQL = "UPDATE organizations SET donotknockexpiration = '" & dbsafe(request.form("years")) & "' WHERE orgid = '" & session("orgid") & "'"
	RunSQLStatement(sSQL)
end if
%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
  Dim sError

 'Set Timezone information into session
  session("iUserOffset") = request.cookies("tz")

 'Override of value from common.asp
sLevel = "../" ' Override of value from common.asp
sView = ""



%>
<html>
<head>
  <title><%=langBSHome%></title>

  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
  <script language="javascript" src="../scripts/ajaxLib.js"></script>

</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content" style="width:auto;">
 	<div id="centercontent">

<table id="bodytable" border="0" cellpadding="0" cellspacing="0" class="start">
  <tr valign="top">
    	<td>
	<h2>Do Not Knock List</h2>

	<%
 	lcl_orghasfeature_expiration                    = orghasfeature("donotknockexpire")

	if lcl_orghasfeature_expiration then
		intYears = 0
		sSQL = "SELECT donotknockexpiration FROM organizations WHERE orgid = '" & session("orgid") & "'"
		Set oE = Server.CreateObject("ADODB.RecordSet")
		oE.Open sSQL, Application("DSN"), 3, 1
		if not oE.EOF then intYears = oE("donotknockexpiration")
		oE.Close
		Set oE = Nothing
	%>
	<script>
		function checkVal()
		{
			var value = document.dnnexp.years.value;

			if (!isNaN(value) && (function(x) { return (x | 0) === x; })(parseFloat(value)))
			{
				document.dnnexp.submit();
			}
			else
			{
				alert("You must enter a whole number of years.");
			}
		}
	</script>
	<form action="#" name="dnnexp" method="POST">
		Registration is active for <input type="text" size="3" name="years" value="<%=intYears%>" /> years <input type="button" value="Save" onClick="checkVal();" />
	</form><br />

	<%
	end if

	sSQL = BuildQuery("solicitors")
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1

	blnSolicitors = false
	if not oRs.EOF then 
		blnSolicitors = true

		response.write "<h3>Do Not Knock - Solicitors <input type=""button"" value=""export"" onclick=""window.location='donotknocklist.asp?export=solicitors';"" /></h3><br />"
		ListMembers oRs
	end if
	oRs.Close

	sSQL = BuildQuery("peddlers")
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1

	if not oRs.EOF then 
		if blnSolicitors then response.write "<br /><br /><br /><br />"
		response.write "<h3>Do Not Knock - Peddlers <input type=""button"" value=""export"" onclick=""window.location='donotknocklist.asp?export=peddlers';"" /></h3><br />"
		ListMembers oRs
	end if
	oRs.Close


	Set oRs = Nothing
	%>
      </td>
  </tr>
</table>

  </div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
Function BuildQuery(strType)
	BuildQuery = "SELECT userfname,userlname,useraddress,useraddress2,usercity,userstate,userzip " _
  		& " FROM egov_users u " _
		& " INNER JOIN Organizations o ON o.OrgID = u.orgid " _
  		& " WHERE  isOnDoNotKnockList_" & strType & " = 1 and u.isdeleted = 0 and u.orgid = " & session("orgid") _
		& " AND (o.donotknockexpiration IS NULL or o.donotknockexpiration = 0 OR DATEADD(yyyy,o.donotknockexpiration,u.donotknockregdate) > GETDATE()) " _
		& " ORDER BY userlname, userfname, useraddress, useraddress2, usercity, userstate, userzip"

End Function


sub ListMembers(oUsers)
	response.write "<table class=""tablelist"" cellspacing=""0"" cellpadding=""2"" border=""0"" style=""min-width:1000px"">"
	response.write "<tr><th>Last Name</th><th>First Name</th><th>Address</th></tr>"
	i = 0
	Do While Not oUsers.EOF
		i = i + 1
		EventOrNot=(i+2) Mod 2
		If EventOrNot = 0 Then 
			sRowClass = ""
		Else 
			sRowClass = " class=""altrow"" "
		End If 
		response.write "<tr " & sRowClass & ">"
		response.write "<td>" & oUsers("userlname") & "</td>"
		response.write "<td>" & oUsers("userfname") & "</td>"
		response.write "<td>" 
				response.write oUsers("useraddress") & "<br />"
				if oUsers("useraddress2") <> "" then response.write oUsers("useraddress2") & "<br />"
				response.write oUsers("usercity") & ", "
				response.write oUsers("userstate") & " "
				response.write oUsers("userzip") 
		response.write "</td>" 
		response.write "</tr>"
		oUsers.MoveNext
	loop
	response.write "</table>"
End Sub
%>
