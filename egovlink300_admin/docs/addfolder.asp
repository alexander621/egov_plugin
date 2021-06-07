<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addfolder.asp
' AUTHOR: Steve Loar
' CREATED: 08/30/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Add folders for documents
'
' MODIFICATION HISTORY
' 1.0   08/30/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sParentFolder, sFolderName, sLoadMsg

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage documents", sLevel	' In common.asp

sParentFolder = request("path")

If request("foldername") <> "" Then
	sFolderName = request("foldername")
Else
	sFolderName = ""
End If 

sSuccessFlag = request("sf")
If sSuccessFlag = "fa" Then
	sLoadMsg = "displayScreenMsg('The folder, " & sFolderName & ", has been successfully added.');"
End If 
If sSuccessFlag = "fe" Then
	sLoadMsg = "displayScreenMsg('The folder, " & sFolderName & ", already exists. Please check the name and try again.');"
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

		function ValidateFolderName()
		{
			//alert($("#foldername").val());

			if ($("#foldername").val() == "")
			{
				alert("Please enter a folder name");
				$("#foldername").focus();
			}
			else
			{
				var filename = $("#foldername").val();

				var rege = /^[\w- :\\]+$/;
				var Ok = rege.test(filename);

				if (! Ok)
				{
					alert ("The folder name has characters that are not allowed. Allowed characters include [A-Za-z0-9_-] and spaces\n\n Example: Documents and Settings \n\n Please rename the folder, removing any special characters from the folder name.");
					$("#foldername").focus();
				}
				else
				{
					document.frmAddFolder.submit();
				}
			}
		}

		$(document).ready( function() {
			$("#foldername").focus();
			//loader();
		});

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
				<font size="+1"><strong>Documents: Add Folder</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
				<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='default.asp'>Back To Documents</a>
			</td></tr></table>

			<form name="frmAddFolder" action="addfolderdo.asp" method="post">
				<input type="hidden" name="path" value="<%=sParentFolder%>" />

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

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>

