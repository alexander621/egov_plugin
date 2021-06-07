<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: deletefolder.asp
' AUTHOR: Steve Loar
' CREATED: 08/31/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Delete Folders
'
' MODIFICATION HISTORY
' 1.0   08/31/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sTargetFolder, sFolderName, iFolderId, sDBFolder

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage documents", sLevel	' In common.asp

sTargetFolder = dbsafe(request("path"))

If Right(sTargetFolder,1) = "/" Then
	' strip off the trailing "/"
	sTargetFolder = Left(sTargetFolder,(Len(sTargetFolder) - 1))
End If 

'sDBFolder = Replace(sTargetFolder, "egovlink300_docs", "public_documents300")
sDBFolder = Replace(sTargetFolder, Application("DocumentsRootDirectory"), "public_documents300")

If request("folderid") <> "" Then
	iFolderId = request("folderid")
Else 
	iFolderId = GetFolderId( sDBFolder )
End If 

If request("foldername") <> "" then
	sFolderName = request("foldername")
Else
	sFolderName = GetFolderName( iFolderId )
End If 
'response.write sFolderName & "<br /><br />"

sSuccessFlag = request("sf")
If sSuccessFlag = "nf" Then
	sLoadMsg = "displayScreenMsg('The folder you attempted to delete, " & sFolderName & ", does not exist. Please try again.');"
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

			if ($("#deletefolderbutton").attr('disabled'))
				$("#deletefolderbutton").attr('disabled','');
			else
				$("#deletefolderbutton").attr('disabled','disabled');
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
				<font size="+1"><strong>Documents: Delete Folder</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
				<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='default.asp'>Back To Documents</a>
			</td></tr></table>

			<form name="frmDeleteFolder" action="deletefolderdo.asp" method="post">
				<input type="hidden" name="folderid" value="<%=iFolderId%>" />
				<input type="hidden" name="path" value="<%=sTargetFolder%>" />
				<input type="hidden" name="foldername" value="<%=sFolderName%>" />

				<p id="folderpath">
					<strong>Folder:</strong> <%=sTargetFolder%>
				</p>

				<p id="actionfield">
					<input type="checkbox" name="DeleteValidate" onclick="ToggleButton();" /> &nbsp; 
					<strong>Confirm the deletion of the &quot;<%=sFolderName%>&quot; folder and all below it?</strong>
				</p>

				<p id="buttons">
					<input type="submit" class="button" id="deletefolderbutton" value="Delete Folder" disabled="disabled" /> &nbsp; &nbsp;
					<input type="button" class="button" value="Cancel" onclick="location.href='default.asp';" />
				</p>

			</form>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


