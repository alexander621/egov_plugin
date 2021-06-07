<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
 response.write "Create Memberids did not run.  PLEASE DISABLE RESPONSE.END TO RUN SCRIPT."
 response.end

'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: create_memberids.asp
' AUTHOR: Steve Loar
' CREATED: 06/19/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module created member ids
'
' MODIFICATION HISTORY
' 1.0   06/19/07	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' USER VALUES

sLevel = "../" ' Override of value from common.asp


%>

<html>
<head>
	<title>E-Gov Memberid creation</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<h1>E-Gov Memberid creation</h1>
	<p><strong>Started: <%=Now()%></strong></p>
	<p><hr /></p>

<%
	Dim sSql, oRs, iMemberid
	iMemberid = 0

	' Update egov_verisign_payment_information - Public Web Purchases
	sSql = "select poolpassid, familymemberid from egov_poolpassmembers order by poolpassid, familymemberid"
	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.Eof
		iMemberid = iMemberid + 1
		sSql = "Update egov_poolpassmembers Set memberid = " & iMemberid & " Where poolpassid = " & oRs("poolpassid") & " and familymemberid = " & oRs("familymemberid")
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