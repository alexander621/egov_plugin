<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
	response.write "Actionline Update did not run.  PLEASE DISABLE RESPONSE.END TO RUN SCRIPT."
	response.end

'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: acriopnline_update.asp
' AUTHOR: Steve Loar
' CREATED: 07/18/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module update actionline data
'
' MODIFICATION HISTORY
' 1.0   07/18/07	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' USER VALUES

sLevel = "../" ' Override of value from common.asp


%>

<html>
<head>
	<title>E-Gov Classes Migration Script</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<h1>E-Gov Actionline Update Script</h1>
	<p><strong>Started: <%=Now()%></strong></p>
	<p><hr /></p>
<%
	sSql = "SELECT action_autoid, comment from egov_actionline_requests where action_autoid > 19235 order by action_autoid"

	response.write "<p>" & sSQL & "</p><br /><br />"
	response.flush

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.Eof
		sComment = Replace(oRs("comment"),"<br />","<br>")
		sSql = "Update egov_actionline_requests Set comment = '" & dbsafe(sComment) & "' Where action_autoid = " & oRs("action_autoid")
		RunSQL( sSql )
		oRs.MoveNext
	Loop
	oRs.Close
	Set oRs = Nothing 

%>

	<p><hr /></p>
	<p><strong>Finished: <%=Now()%></strong></p>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( sSql )
	Dim oCmd

	response.write "<p>" & sSql & "</p><br /><br />"
	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 

%>