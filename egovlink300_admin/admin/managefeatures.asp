<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="feature_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: manage_features.asp
' AUTHOR: Steve Loar
' CREATED: 09/19/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the management of features for clients
'
' MODIFICATION HISTORY
' 1.0	09/12/2008	Steve Loar - Changed to just manage features. See featureselection for setting to orgs
' 1.1	04/17/2009	David Boyer - Added "CommunityLinkOn" column
' 1.2	07/27/2009	David Boyer - Added screen messages
' 1.3	10/20/2011	Steve Loar - Converted to HTML 5, variables given proper scope, and re-formatted as I created it.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim lcl_onload, lcl_success, lcl_msg

sLevel = "../"  'Override of value from common.asp

If Not UserIsRootAdmin(session("userid")) Then 
	response.redirect "../default.asp"
End If 

'Check for a screen message
lcl_onload  = ""
lcl_success = request("success")

If lcl_success <> "" Then 
	lcl_msg = setupScreenMsg(lcl_success)
	lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
End If 


%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-Gov Administration Console {Feature Management}</title>
	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="javascript">
	<!--
		function displayScreenMsg( iMsg ) 
		{
			if(iMsg!="") 
			{
				document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg( ) 
		{
			document.getElementById("screenMsg").innerHTML = "";
		}

	//-->
	</script>

</head>
<body onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<table border="0" cellspacing="0" cellpadding="0" style="width:1100px;">
		<tr><td colspan="2"><h3>Feature Management</h3></td></tr>
		<tr>
			<td>
				<input type="button" class="button" name="back" value="<< Back to Feature Selection" onClick="javascript:window.location='featureselection.asp?orgid=<%=session("orgid")%>'" />
				<input type="button" class="button" name="create" value="New Feature" onClick="javascript:window.location='featureedit.asp?featureid=0'" />
				<input type="button" class="button" name="order" value="Order Features" onClick="javascript:window.location='featureordering.asp'" />
			</td>
			<td width="40%" align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
		</tr>
	</table>

	<table id="featuremanagelist" border="0" cellpadding="3" cellspacing="0">
		<tr>
			<th>Feature</th>
			<th>FeatureId</th>
			<th>ParentId</th>
			<th>Feature<br />Type</th>
			<th align="left">FeatureName</th>
			<th width="200" nowrap="nowrap">Feature Notes</th>
			<th align="center">Avail on<br />Community Link</th>
		</tr>
		<% ShowFeatures 0 %>
	</table>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%

'------------------------------------------------------------------------------
' void ShowFeatures iParentFeatureId
'------------------------------------------------------------------------------
Sub ShowFeatures( ByVal iParentFeatureId )
	Dim sSql, oRs, iRows, sFeatureNotes, sFeatureDescription

	iRows = 0
	sSql = "SELECT featureid, feature, featurename, featuretype, parentfeatureid, ISNULL(featurenotes,'') AS featurenotes, CommunityLinkOn "
	sSql = sSql & " FROM egov_organization_features "
	sSql = sSql & " WHERE parentfeatureid = " & iParentFeatureId
	
	If iParentFeatureId = 0 Then 
		sSql = sSql & " ORDER BY admindisplayorder"
	Else 
		sSql = sSql & " ORDER BY securitydisplayorder"
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iRows = iRows + 1
		lcl_td_onclick = "location.href='featureedit.asp?featureid=" & oRs("featureid") & "';"

		response.write vbcrlf & "<tr"

		If iParentFeatureId = 0 Then 
			response.write " class=""parentfeature"" "
		Else 
			If iRows Mod 2 = 0 Then 
				response.write " class=""altrow"" "
			End If 
		End If 

		response.write " onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">"
		response.write "<td nowrap=""nowrap"" valign=""top"" onClick=""" & lcl_td_onclick & """>" 

		If iParentFeatureId > 0 Then 
			response.write " &nbsp; &bull; " & oRs("featurename")
		Else 
			response.write "<strong>" & oRs("featurename") & " </strong>"
		End If 

		response.write "</td>"
		response.write "<td align=""center"" valign=""top"" onClick=""" & lcl_td_onclick & """>" & oRs("featureid") & "</td>"
		response.write "<td align=""center"" valign=""top"" onClick=""" & lcl_td_onclick & """>" & oRs("parentfeatureid") & "</td>"
		response.write "<td align=""center"" valign=""top"" onClick=""" & lcl_td_onclick & """>"

		If oRs("featuretype") = "N" Then 
			response.write "Navigation"
		Else
			response.write "Security"
		End If 
		response.write "</td>"

		response.write "<td valign=""top"">" & oRs("feature") & "</td>"

		response.write "<td valign=""top"" onClick=""" & lcl_td_onclick & """>" 

		sFeatureNotes = ""
		sFeatureNotes = CStr(oRs("featurenotes"))

		If Trim(sFeatureNotes) = "" Then 
			response.write "&nbsp;"
		Else 
			response.write sFeatureNotes
		End If 

		response.write "</td>"

	   'Setup the display values for CommunityLinkOn
		If oRs("CommunityLinkOn") Then 
		   lcl_communityLinkOn = "Y"
		Else 
		   lcl_communityLinkOn = "&nbsp"
		End If 

		response.write "<td align=""center"" valign=""top"" onClick=""" & lcl_td_onclick & """>" & lcl_communityLinkOn & "</td>"
		response.write "</tr>"

		ShowFeatures oRs("featureid")
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing
	
End Sub 



%>