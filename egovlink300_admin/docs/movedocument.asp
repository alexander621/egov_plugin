<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: movedocument.asp
' AUTHOR: Steve Loar
' CREATED: 09/08/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Moves a Document
'
' MODIFICATION HISTORY
' 1.0   09/08/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sTargetFile, sDBFile, iFileId, sFileName, iFolderId, sNewDBFolder

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

sNewDBFolder = Left(sDBFile,Len(sDBFile) - (Len(sFileName)+1))

If request("folderid") <> "" Then
	iFolderId = request("folderid")
Else 
	iFolderId = GetFolderId( sNewDBFolder )
End If 


sSuccessFlag = request("sf")
If sSuccessFlag = "df" Then
	sLoadMsg = "displayScreenMsg('A document by that name already exists in the selected folder. Please try again.');"
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

		function MoveFile()
		{
			document.frmMoveFile.submit();
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
				<font size="+1"><strong>Documents: Move Document</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
				<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='default.asp'>Back To Documents</a>
			</td></tr></table>

			<form name="frmMoveFile" action="movedocumentdo.asp" method="post">
				<input type="hidden" name="fileid" value="<%=iFileId%>" />
				<input type="hidden" name="path" value="<%=sTargetFile%>" />
				<input type="hidden" name="filename" value="<%=sFileName%>" />

				<p id="folderpath">
					<strong>Document:</strong> <%=sTargetFile%>
				</p>
				<p id="actionfield">
					<strong>New Folder:</strong><br />
					<% ShowFolders iFolderId %>
				</p>

				<p id="buttons">
					<input type="button" class="button" value="Move Document" onclick="MoveFile();" /> &nbsp; &nbsp;
					<input type="button" class="button" value="Cancel" onclick="location.href='default.asp';" />
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
' void ShowFolders iFolderId 
'--------------------------------------------------------------------------------------------------
Sub ShowFolders( ByVal iFolderId )
	Dim sSql, oRs

	sSql = "SELECT folderid, folderpath FROM documentfolders "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY folderpath"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""folderid"" id=""folderid"">"
		Do While Not oRs.EOF
			' /public_documents300/custom/pub/
			sDocPath = Replace(oRs("folderpath"),"/public_documents300/custom/pub/" & session("virtualdirectory"),"")
			If sDocPath <> "" And sDocPath <> "/published_documents" And sDocPath <> "/unpublished_documents" Then 
				response.write vbcrlf & "<option value=""" & oRs("folderid") & """"
				If CLng(iFolderId) = CLng(oRs("folderid")) Then
					response.write " selected=""selected"" "
				End If 
				response.write ">" & Mid(sDocPath,2) & "</option>"
			End If 
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>