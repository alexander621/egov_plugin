<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: mergecitizensresult.asp
' AUTHOR: Steve Loar
' CREATED: 1/5/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Merge Citizen records together. Merges the entire family.
'
' MODIFICATION HISTORY
' 1.0   1/5/2009	Steve Loar - INITIAL VERSION Started
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFamilyId, sMergeName, sKeepName, iMergeUserId, iKeepUserId

sLevel = "../" ' Override of value from common.asp

' Check the page availability and user access rights in one call
PageDisplayCheck "merge registered users", sLevel	' In common.asp

iFamilyId = CLng(request("familyid"))

iMergeUserId = CLng(request("mergeuserid"))
sMergeName = GetCitizenName( iMergeUserId )		' In common.asp

iKeepUserId = CLng(request("keepuserid"))
sKeepName = GetCitizenName( iKeepUserId )		' in common.asp

Session("RedirectLang") = "Edit Citizen Users"
Session("RedirectPage") = "display_citizen.asp"
%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="mergecitizens.css" />

</head>
<body>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Merge Completed</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<p>
				The merge of <%=sMergeName%> into <%=sKeepName%> was successful.
			</p>
			<p>
				<input type="button" class="button" value="Merge Another" onclick="javascript:window.location='mergecitizens.asp';" /> &nbsp;&nbsp;
<%				If UserHasPermission( Session("UserId"), "edit citizens" ) Then		%>				
					<input type="button" class="button" value="View Resulting Household" onclick="javascript:window.location='family_list.asp?userid=<%=iFamilyId%>';" /> 
<%				End If		%>
			</p>

		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

