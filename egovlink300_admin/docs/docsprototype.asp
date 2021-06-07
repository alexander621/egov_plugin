<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: editsecurity.asp
' AUTHOR: Steve Loar
' CREATED: 08/31/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Documents Prototype page
'
' MODIFICATION HISTORY
' 1.0   08/31/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage documents", sLevel	' In common.asp


sSuccessFlag = request("sf")
If sSuccessFlag = "rc" Then
	sLoadMsg = "displayScreenMsg('The reservation has been successfully made.');"
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="docstyles.css" />

	<script type="text/javascript" src="../scripts/jquery-1.4.2.min.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="Javascript">
	<!--

		function loader()
		{
			<%=sLoadMsg%>
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html("*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;");
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("&nbsp;");
		}


	//-->
	</script>

</head>

<body onload="loader();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Documents: Edit Security</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
			</td></tr></table>

			<form name="frmAddFolder" action="addfolderdo.asp" method="post">
				<p id="folderpath">
					<strong>Parent Folder:</strong> <%=sParentFolder%>
				</p>
				<p id="actionfield">
					<strong>New Folder Name:</strong> <input type="text" name="foldername" id="foldername" value="<%=sFolderName%>" size="100" maxlength="100" />
				</p>

				<p id="buttons">
					<input type="button" class="button" value="Add New Folder" onclick="ValidateFolderName();" /> &nbsp; &nbsp;
					<input type="button" class="button" value="Cancel" onclick="location.href='default.asp';" />
				</p>

			</form>


<%				'Pull the list here
%>			

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>