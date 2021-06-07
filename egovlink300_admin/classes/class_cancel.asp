<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_cancel.asp
' AUTHOR: Steve Loar
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page is for cancelling a class.
'
' MODIFICATION HISTORY
' 1.0   04/26/06   Steve Loar - INITIAL VERSION
' 1.1	08/20/07   Steve Loar - Changed to new menu navigation
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iClassId

iClassId = CLng(request("classid"))

sLevel = "../" ' Override of value from common.asp

%>

<html>
<head>
	<title>E-Gov Administration Console</title>
	
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

<script language="Javascript">
<!--

	function Validate() 
	{
		if (document.CancelForm.cancelreason.value == "")
		{
			alert('Please provide a reason for cancelling.');
			document.CancelForm.cancelreason.focus();
			return;
		}

		document.CancelForm.submit();
	}

//-->
</script>

</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		<h3>Cancel <%=GetClassName( iClassId )%></h3>

		<form name="CancelForm" method="post" action="class_changestatus.asp">
		<input type="hidden" name="classid" value="<%=iClassId%>" />
		<input type="hidden" name="statusid" value="2" />
			<p>
			<strong>Reason for Cancelling</strong><br />
			<textarea name="cancelreason" id="classdescription"></textarea>
			</p>
			<p>
				<input type="checkbox" name="emailclass" /> Send notification email to attendees
			</p>
			<p>
				<input type="button" class="button" name="cancel" value="Cancel Class/Event" onclick="Validate();" /> &nbsp; &nbsp; 
				<input type="button" class="button" name="return" value="Return to Edit" onclick="javascript:location.href='edit_class.asp?classid=<%=iClassId%>';" />
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


