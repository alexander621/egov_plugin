<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: editpublicaccess.asp
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
Dim sSuccessFlag, sTargetFolder, sDBFolder, iFolderId

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "manage documents", sLevel	' In common.asp

sTargetFolder = dbsafe(request("path"))

'sDBFolder = Replace(sTargetFolder, "egovlink300_docs", "public_documents300")
sDBFolder = Replace(sTargetFolder, Application("DocumentsRootDirectory"), "public_documents300")

sDBFolder = Left(sDBFolder, Len(sDBFolder)-1) ' want to remove the ending "/"

iFolderId = GetFolderId( sDBFolder )

sSuccessFlag = request("sf")
If sSuccessFlag = "pc" Then
	sLoadMsg = "displayScreenMsg('Permissions successfully changed.');"
End If 

bBack = HasExistingMembers( iFolderId )
bForward = HasAvailableMembers( iFolderId )

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
				<font size="+1"><strong>Documents: Edit Public Access</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg">&nbsp;</span>
				<img src='../images/arrow_2back.gif' align='absmiddle'>&nbsp;<a href='default.asp'>Back To Documents</a>
			</td></tr></table>

			<p id="folderpath">
				<strong>Folder:</strong> <%=sDBFolder%>
			</p><br /><br />

			<table border="0" cellpadding="5" cellspacing="0" id="publicaccesscontainer">
				<tr>
					<td>
						<form name="c1" method="post" action="editpublicaccessdo.asp">
							<input type="hidden" name="t" value="del" />
							<input type="hidden" name="path" value="<%=sDBFolder%>" />
							<input type="hidden" name="folderid" value="<%=iFolderId%>" />
							<table border="0" cellpadding="0" cellspacing="0" width="130">
								<tr>
									<td><b><%= langMember & "&nbsp;" & langCommittees %></b></td>
								</tr>
								<tr><td><br></td></tr>
								<tr>
									<td><% ShowExistingMembers sDBFolder %></td>
								</tr>
							</table>
							<input type="hidden" name="sMsg" value="Permissions successfully changed.">
						</form>
					</td>
					<td>
						<% If bBack Then %>
							<a href="javascript:document.c1.submit();"><img src="../images/ieforward.gif" border="0" /></a><br />
						<% Else %>
							<img src="../images/ieforward_disabled.gif" border="0" /></a><br />
						<% End If %>
						<br />
						<% If bForward Then %>
							<a href="javascript:document.r1.submit();"><img src="../images/ieback.gif" border="0" /></a>
						<% Else %>
							<img src="../images/ieback_disabled.gif" border="0" /></a>
						<% End If %>
					</td>
					<td>
						<form name="r1" method="post" action="editpublicaccessdo.asp">
							<input type="hidden" name="t" value="add" />
							<input type="hidden" name="path" value="<%=sDBFolder%>" />
							<input type="hidden" name="folderid" value="<%=iFolderId%>" />
							<table border="0" cellpadding="0" cellspacing="0" width="130">
								<tr>
									<td><b><%=langCommittee%> List</b></td>
								</tr>
								<tr><td><br></td></tr>
								<tr>
									<td><% GetAvailableMembers sDBFolder %></td>
								</tr>
							</table>
							<input type="hidden" name="sMsg" value="Permissions successfully changed." id="Hidden1">
						</form>
					</td>
				</tr>
				<tr><td class="success"><%=request("sMSG")%></td></tr>
			</table>

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

'--------------------------------------------------------------------------------------------------
' void ShowExistingMembers sPath
'--------------------------------------------------------------------------------------------------
Sub ShowExistingMembers( ByVal sPath )
	Dim sSql, oRs

	sSql = "SELECT g.GroupID, g.GroupName "
	sSql = sSql & "FROM DocumentFolders [df] INNER JOIN CitizenFeatureAccess [fa] ON fa.AccessID = df.CitizenAccessID "
	sSql = sSql & "INNER JOIN citizenGroups [g] ON g.GroupID = fa.GroupID "
	sSql = sSql & "WHERE fa.OrgID = " & Session("OrgID") & " AND df.FolderPath = '" & sPath & "' "
	sSql = sSql & "ORDER BY GroupName"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select size=""20"" style=""width:200px"" name=""ExistingList"" multiple>"
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=" & oRs("GroupID") & ">" & oRs("GroupName") & "</option>"
			oRs.MoveNext
		Loop 
	Else 
		response.write vbcrlf & "<option value=""-1"">Everyone</option>"
	End If

	response.write vbcrlf & "</select>"  

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean HasExistingMembers( iFolderId )
'--------------------------------------------------------------------------------------------------
Function HasExistingMembers( ByVal iFolderId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(g.GroupID) AS hits "
	sSql = sSql & "FROM DocumentFolders [df] INNER JOIN CitizenFeatureAccess [fa] ON fa.AccessID = df.CitizenAccessID "
	sSql = sSql & "INNER JOIN citizenGroups [g] ON g.GroupID = fa.GroupID "
	sSql = sSql & "WHERE fa.OrgID = " & Session("OrgID") & " AND df.Folderid = " & iFolderId & " "
	'response.write sSql & "<br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			HasExistingMembers = True 
		Else
			HasExistingMembers = False 
		End If 
	Else 
		HasExistingMembers = False 
	End If

	oRs.Close
	Set oRs = Nothing 

End Function  


'--------------------------------------------------------------------------------------------------
' void GetExistingMembers sPath
'--------------------------------------------------------------------------------------------------
Sub GetAvailableMembers( ByVal sPath )
	Dim sSql, oRs

	sSql = "SELECT g.GroupID, g.GroupName "
	sSql = sSql & "FROM citizenGroups [g] WHERE OrgID = " & Session("OrgID") & " AND GroupID NOT IN "
	sSql = sSql & "(SELECT g2.GroupID FROM DocumentFolders [df] INNER JOIN CitizenFeatureAccess [fa] ON "
	sSql = sSql & "fa.AccessID = df.CitizenAccessID INNER JOIN citizenGroups [g2] ON g2.GroupID = fa.GroupID "
	sSql = sSql & "WHERE df.FolderPath = '" & sPath & "') "
	sSql = sSql & "ORDER BY GroupName"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select size=""20"" style=""width:200px;"" name=""RemainingList"" multiple>"

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=" & oRs("GroupID") & ">" & oRs("GroupName") & "</option>"
		oRs.MoveNext
	Loop

	response.write vbcrlf & "</select>" 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean HasAvailableMembers( iFolderId )
'--------------------------------------------------------------------------------------------------
Function HasAvailableMembers( ByVal iFolderId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(g.GroupID) AS hits "
	sSql = sSql & "FROM citizenGroups [g] WHERE OrgID = " & Session("OrgID") & " AND GroupID NOT IN "
	sSql = sSql & "(SELECT g2.GroupID FROM DocumentFolders [df] INNER JOIN CitizenFeatureAccess [fa] ON "
	sSql = sSql & "fa.AccessID = df.CitizenAccessID INNER JOIN citizenGroups [g2] ON g2.GroupID = fa.GroupID "
	sSql = sSql & "WHERE df.FolderId = " & iFolderId & ") "
	'response.write sSql & "<br />"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			HasAvailableMembers = True 
		Else
			HasAvailableMembers = False 
		End If 
	Else 
		HasAvailableMembers = False 
	End If

	oRs.Close
	Set oRs = Nothing 

End Function  



%>