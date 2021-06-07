<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<%
	sSQL = "SELECT u.UserID, FirstName, MiddleInitial, LastName, BusinessNumber, u.Email, JobTitle, " _
		& " CASE WHEN Department IS NULL OR Department = '' THEN g.org_name ELSE Department END as Department " _
		& " FROM users u " _
		& " INNER JOIN egov_staff_directory_usergroups ug ON ug.userid = u.UserID " _
		& " LEFT JOIN egov_staff_directory_groups g ON g.org_group_id = ug.org_group_id " _
		& " WHERE u.orgid = '" & iorgid & "'"
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
Do While Not oRs.EOF
	response.write "dn: uid=" & oRs("userid") & ",ou=People,dc=egovlink,dc=com " & vbcrlf
	response.write "objectClass: person " & vbcrlf
	response.write "uid: " & oRs("userid") & vbcrlf
	response.write "givenname:  " & oRs("FirstName") & vbcrlf
	response.write "cn: " & oRs("FirstName") & oRs("MiddleInitial") & oRs("LastName") & vbcrlf
	response.write "telephonenumber: " & oRs("BusinessNumber") & vbcrlf
	response.write "sn: " & oRs("lastname") & vbcrlf
	response.write "mail: " & oRs("email") & vbcrlf
	response.write "title: " & oRs("jobtitle") & vbcrlf
	response.write "ou: " & oRs("Department") & vbcrlf
	response.write "ou: People " & vbcrlf
	response.write vbcrlf
	oRs.MoveNext
loop
oRs.Close
Set oRs = Nothing
%>
