<%
Dim iMaillistid,iMaillistname
iMaillistid = request("idlid")
iMaillistname = request("iname")
%>

<html>

<head>
<link href="../global.css" rel="stylesheet" type="text/css">
<link href="classes.css" rel="stylesheet" type="text/css">
</head>

<body  bgcolor="#c9def0">


	<table border="0" cellpadding="10" cellspacing="0" width="100%" bgcolor="#c9def0">
	    <tr>
	      <td colspan="2" valign="top">
  
				<%
				' DISPLAY LIST OF SUBSCRIBED MAILING LISTS
				response.write "<center>"
				response.write "<strong>Distribution List: "& iMaillistname &" </strong>"
				response.write "<br /><a href='javascript:self.close();'><font size=2>Close Window</font></a></font></center>"
				response.write "<table cellpadding=10 cellspacing=0 width=350 border=0>"
				response.write "<tr><td>"
					subDisplaySubscribedUsers
				response.write "</td>"

				' DISPLAY ARROWS
				response.write "<td align=center>"
				response.write "&nbsp;<a href='javascript:document.sl.submit();'><img src='../images/ieforward.gif' align='absmiddle' border=0></a>"
				response.write "<br><br>"
				response.write "<a href='javascript:document.al.submit();'><img src='../images/ieback.gif' align='absmiddle' border=0></a>"
				response.write "</td>"

				' DISPLAY LIST OF AVAILABLE MAILING LISTS
				response.write "<td>"
					subDisplayAvailableUsers
				response.write "</td>"
				response.write "</tr>"
				response.write "</table>"
				%>

		  </td>
		</tr>
	</table>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYSUBSCRIBEDUSERS
'--------------------------------------------------------------------------------------------------
Sub subDisplaySubscribedUsers
	
	sSQL = "SELECT * FROM egov_users u INNER JOIN egov_class_distributionlist_to_user ug ON u.userid=ug.userid where (ug.distributionlistid = '" & iMaillistid & "') ORDER BY u.userlname "
	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write "<table border=0 cellpadding=0 cellspacing=0 width=130 >"
	response.write "<tr><td height=20><b>SUBSCRIBED</b></td></tr>"
	response.write "<tr><td>"
	response.write "<form name=sl method=""post"" action='dl_deletemember.asp'>"
	response.write "<input type=hidden name=maillistid value=""" & iMaillistid & """>"
	response.write "<input type=hidden name=maillistname value=""" & iMaillistname & """>"
	response.write "<select size='15' border=0 width=""140"" STYLE=""width:140px"" name='subscribedlist' multiple>"

	' LOOP THRU AVAILABLE LISTS
	If NOT oList.EOF Then
		Do While Not oList.EOF
			response.write "<option value=""" & oList("userid") & """>" & oList("userlname") & ", " & oList("userfname")
			oList.MoveNext
		Loop
	End If

	response.write   "</select></p>"
	response.write  "</form>"
	response.write "</td></tr></table>"  

end sub


'--------------------------------------------------------------------------------------------------
' SUB SUBDISPLAYAVAILABLEUSERS
'--------------------------------------------------------------------------------------------------
Sub SubDisplayAvailableUsers

	sSQL = "SELECT * FROM egov_users AS u WHERE userregistered = 1 and useremail is not NULL and userlname is not NULL and userfname is not NULL and "
	sSql = sSql & " (userid NOT IN (SELECT userid FROM egov_class_distributionlist_to_user AS ug "
	sSql = sSql & " WHERE (distributionlistid = '" & iMaillistid & "'))) AND (orgid = '" & SESSION("ORGID") & "') ORDER BY userlname "

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write "<table border=0 cellpadding=0 cellspacing=0 width=130 >"
	response.write "<tr><td height=20><b>AVAILABLE</b></td></tr>"
	response.write "<tr><td>"
	response.write "<form name=al method=""post"" action='dl_addmember.asp'>"
	response.write "<input type=hidden name=maillistid value=""" & iMaillistid & """>"
	response.write "<input type=hidden name=maillistname value=""" & iMaillistname & """>"
	response.write "<select size='15' border=0 WIDTH=""140"" STYLE=""width:140px"" name='availablelist' multiple>"

	' LOOP THRU AVAILABLE LISTS
	If NOT oList.EOF Then
		Do While Not oList.EOF
			response.write "<option value=""" & oList("userid") & """>" & oList("userlname") & ", " & oList("userfname")
			oList.MoveNext
		Loop
	End If

	response.write   "</select></p>"
	response.write  "</form>"
	response.write "</td></tr></table>"  
End Sub

%>

</body>

</html>