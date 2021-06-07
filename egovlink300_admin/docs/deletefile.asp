<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: deletefile.asp
' AUTHOR: Steve Loar
' CREATED: 09/03/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Document delete page
'
' MODIFICATION HISTORY
' 1.0   09/03/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sTargetFile, iFileId, sFileName, sDBFile

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage documents", sLevel	' In common.asp

sTargetFile = dbsafe(request("path"))

sDBFile = Replace(sTargetFile, Application("DocumentsRootDirectory"), "public_documents300")

If request("folderid") <> "" Then
	iFileId = request("fileid")
Else 
	iFileId = GetFileId( sDBFile )
End If 

If request("filename") <> "" then
	sFileName = request("filename")
Else
	sFileName = GetFileName( iFileId )
End If 

sSuccessFlag = request("sf")
If sSuccessFlag = "nf" Then
	sLoadMsg = "displayScreenMsg('The document you attempted to delete, " & sFileName & ", does not exist. Please try again.');"
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

		function ToggleButton()
		{
			//alert('Toggle');
			//alert($("#deletefolderbutton").attr('disabled'));

			if ($("#deletefilebutton").attr('disabled'))
				$("#deletefilebutton").attr('disabled','');
			else
				$("#deletefilebutton").attr('disabled','disabled');
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
				<font size="+1"><strong>Documents: Delete a Document</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
				<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='default.asp'>Back To Documents</a>
			</td></tr></table>

			<form name="frmDeleteFolder" action="deletefiledo.asp" method="post">
				<input type="hidden" name="fileid" value="<%=iFileId%>" />
				<input type="hidden" name="path" value="<%=sTargetFile%>" />
				<input type="hidden" name="filename" value="<%=sFileName%>" />

				<p id="folderpath">
					<strong>Document: </strong> <%=sTargetFile%>
				</p>

				<p id="actionfield">
					<input type="checkbox" name="DeleteValidate" onclick="ToggleButton();" /> &nbsp; 
					<strong>Confirm the deletion of this document.</strong>
				</p>

				<p id="buttons">
					<input type="submit" class="button" id="deletefilebutton" value="Delete this Document" disabled="disabled" /> &nbsp; &nbsp;
					<input type="button" class="button" value="Cancel" onclick="location.href='default.asp';" />
				</p>

			</form>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>

