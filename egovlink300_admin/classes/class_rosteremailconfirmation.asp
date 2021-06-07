<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_rosteremailconfirmation.asp
' AUTHOR: Steve Loar
' CREATED: 05/18/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page is for confirms the mass emais to class/event rosters.
'
' MODIFICATION HISTORY
' 1.0   05/18/06   Steve Loar - INITIAL VERSION
' 1.1	10/17/06	Steve Loar - Security, Header and nav changed
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "registration" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

Dim iClassId, iTimeId, iSentCount

iClassId = request("classid")
iTimeId = request("timeid")
iSentCount = request("sentcount")

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

<script language="Javascript">
<!--

//-->
</script>

</head>

<body>
 
<%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		<a href="view_roster.asp?classid=<%=iClassId%>&timeid=<%=iTimeId%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to the Roster</a><br /><br />

		<h3>Email Sent to <%=GetClassName( iClassId ) & " &nbsp; ( " & GetActivityNo( iTimeId ) & " )"%></h3>

		<form name="ConfirmForm" method="post" action="view_roster.asp">
		<input type="hidden" name="classid" value="<%=iClassId%>" />
		<input type="hidden" name="timeid" value="<%=iTimeId%>" />
			<p>
				<%=iSentCount%> messages were sent.
			</p>
			<p>
				<input type="submit" class="button" name="send" value="Return To Roster" /> &nbsp; &nbsp; 
			</p>
		</form>

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


%>


