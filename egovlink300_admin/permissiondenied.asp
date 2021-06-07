<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">

<!-- #include file="includes/common.asp" //-->

<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permissiondenied.ASP
' AUTHOR: Steve Loar
' CREATED: 10/03/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   10/03/06	Steve Loar	- Initial version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = ""

%>

<html>

	<head>
		<title>E-GovLink Administration Console - Permission Denied</title>
		<link rel="stylesheet" type="text/css" href="global.css" />
		<link rel="stylesheet" type="text/css" href="menu/menu_scripts/menu.css" />

	</head>
<body>

<% ShowHeader sLevel %>

<!--#Include file="menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
 <div id="content">
	<div id="centercontent">

	<h3>Permission Denied</h3>

	<p>
		You do not have permissions to the page you are trying to access.  
	</p>
	<p>
		Please contact your administrator if you wish to access this page.
	</p>

</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="admin_footer.asp"-->  

</body>


</html>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>


