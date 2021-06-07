<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: renamedocument.asp
' AUTHOR: Steve Loar
' CREATED: 09/07/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Rename a Document.
'
' MODIFICATION HISTORY
' 1.0   09/07/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sTargetFile, iFileId, sFileName, sDBFile

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage documents", sLevel	' In common.asp


sTargetFile = dbsafe(request("path"))

'sDBFile = Replace(sTargetFile, "egovlink300_docs", "public_documents300")
sDBFile = Replace(sTargetFile, Application("DocumentsRootDirectory"), "public_documents300")

If request("folderid") <> "" Then
	iFileId = request("fileid")
Else 
	iFileId = GetFileId( sDBFile )
End If 

If request("filename") <> "" then
	sFileName = Trim(request("filename"))
Else
	sFileName = Trim(GetFileName( iFileId ))
End If 

sSuccessFlag = request("sf")
If sSuccessFlag <> "" Then 
	If sSuccessFlag = "fe" Then
		sLoadMsg = "displayScreenMsg('A document by that name already exists. Please try again.');"
	End If 
	If sSuccessFlag = "bf" Then
		sLoadMsg = "displayScreenMsg('The new filename has characters that are not allowed. Please try again.');"
	End If 
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

		function displayScreenMsg( iMsg ) 
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

		function ValidateFileName()
		{
			if ($("#NewName").val() == "")
			{

				// If they did not enter a filename, take them back.
				alert("Please enter the new file name.");
				$("#NewName").focus();
			}
			else
			{
				var firstpos;
				var filename = $("#NewName").val();
				var tempname = $("#NewName").val();

				if (tempname.lastIndexOf('\\') != -1)
				{
					firstpos = tempname.lastIndexOf('\\')+1;
					filename = tempname.substring(firstpos);
				}

				//var rege = /^[\w- :\\]+\.{1}[A-Za-z0-9]{2}[A-Za-z0-9]{0,2}$/;
				
				var rege = /^[A-Za-z0-9 _-]+\.{1}[A-Za-z0-9]{2}[A-Za-z0-9]{0,2}$/;
				var Ok = rege.test(filename);

				if (! Ok)
				{
					alert ("The new filename has characters that are not allowed. Allowed characters include [A-Za-z0-9_-], spaces and one '.' \n\n Example: C:\\Documents and Settings\\My Doc_1-2006.txt \n\nPlease remove any special characters and try again.");
					$("#NewName").focus();
				}
				else
				{
					//alert($("#NewName").val());
					document.frmRenameDocument.submit();
				}
			}
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
				<font size="+1"><strong>Documents: Rename Document</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
				<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='default.asp'>Back To Documents</a>
			</td></tr></table>

			<form name="frmRenameDocument" action="renamedocumentdo.asp" method="post">
				<input type="hidden" name="fileid" value="<%=iFileId%>" />
				<input type="hidden" name="path" value="<%=sTargetFile%>" />
				<input type="hidden" name="filename" value="<%=sFileName%>" />

				<p id="folderpath">
					<strong>Document: </strong> <%=sTargetFile%>
				</p>
				<p id="actionfield">
					<strong>New Document Name:</strong> <input type="text" name="NewName" id="NewName" value="<%=sFileName%>" size="100" maxlength="100" />
				</p>

				<p id="buttons">
					<input type="button" class="button" value="Rename Document" onclick="ValidateFileName();" /> &nbsp; &nbsp;
					<input type="button" class="button" value="Cancel" onclick="location.href='default.asp';" />
				</p>

			</form>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>

