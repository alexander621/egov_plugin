<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: manage_features.asp
' AUTHOR: Steve Loar
' CREATED: 09/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the management of features for clients
'
' MODIFICATION HISTORY
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then
	response.redirect "../default.asp"
End If 

If request("orgid") <> "" Then 
	iOrgId = request("orgid")
Else
	iOrgId = GetMaxOrgId
End If 

%>


<html>
<head>
	<title>E-Gov Organization Feature Management</title>
	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />

<script language="Javascript">
<!--

	function ShowOrg()
	{
		document.pickForm.submit();
	}

	function AdminSecurity( sOrgId )
	{
		if (confirm('Run Security Setup?'))
		{
			window.location='admin_security.asp?orgid=' + sOrgId;
		}
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

		<h3>Organization Feature Management</h3>
		
		<form name="pickForm" method="post" action="manage_features.asp">
			<p>
				Organization: <% ShowOrgDropDown iOrgId %>
			</p>
		</form>

		<form name="FeatureForm" method="post" action="save_features.asp">
			<input type="hidden" name="orgid" value="<%=iOrgId%>" />
			<p>
				<input type="submit" class="button" name="save1" value="Save Changes" /> &nbsp; 
				<input type="button" class="button" name="create" value="New Feature" onClick="javascript:window.location='new_feature.asp?orgid=<%=iOrgId%>'" /> &nbsp;
				<input type="button" class="button" name="edit" value="Properties" onClick="javascript:window.location='edit_org.asp?orgid=<%=iOrgId%>'" /> &nbsp; 
				<input type="button" class="button" name="createorg" value="New Org" onClick="javascript:window.location='new_org.asp?orgid=<%=iOrgId%>'" /> &nbsp;
				<input type="button" class="button" name="adminsecurity" value="Initial Admin Security" onClick="AdminSecurity('<%=iOrgId%>')" /> &nbsp;
				<input type="button" class="button" name="orgdisplayedit" value="Edit Displays" onClick="javascript:window.location='org_display_edit.asp?orgid=<%=iOrgId%>'" /><br /><br />
				<input type="button" class="button" name="docsync" value="Document Sync" onClick="javascript:window.location='../docs/docsync.asp?orgid=<%=iOrgId%>'" /> &nbsp;
			</p>
			<p>
				<div class="categories">
					<table id="featuremanager" border="1" cellpadding="3" cellspacing="0">
						<tr><th colspan="2">Feature</th><th>Id</th><th>Type</th><th>Public Feature</th><th>Default Feature</th><th width="250" nowrap="nowrap">Feature Notes</th><th>Public Can View</th><th>Feature Name</th><th>Feature Description</th><th>Public URL</th><th>Public Display Order</th><th>Public Image URL</th>
						</tr>
						<% ShowFeatures iOrgId, 0 %>
					</table>
				</div>
			</p>
			<p>
				<input type="submit" class="button" name="save2" value="Save Changes" />
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


'--------------------------------------------------------------------------------------------------
' Function GetMaxOrgId
'--------------------------------------------------------------------------------------------------
Function GetMaxOrgId
	Dim sSql, oOrgs

	sSql = "Select max(orgid) as maxorgid from organizations"

	Set oOrgs = Server.CreateObject("ADODB.Recordset")
	oOrgs.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	GetMaxOrgId = oOrgs("maxorgid")

	oOrgs.close
	Set oOrgs = Nothing

End Function 

'--------------------------------------------------------------------------------------------------
' Sub ShowOrgDropDown( iOrgId )
'--------------------------------------------------------------------------------------------------
Sub  ShowOrgDropDown( iOrgId )
	Dim sSql, oOrgs

	sSql = "Select orgname, orgcity, orgid, defaultstate from organizations Order By orgcity, defaultstate "

	Set oOrgs = Server.CreateObject("ADODB.Recordset")
	oOrgs.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If Not oOrgs.EOF Then
		response.write vbcrlf & "<select name=""orgid"" onchange='ShowOrg();'>"
		Do While Not oOrgs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oOrgs("orgid") & """ "
			If clng(iOrgId) = clng(oOrgs("orgid")) Then response.write " selected=""selected"" "
			response.write ">" & oOrgs("orgcity") & ", " & oOrgs("defaultstate") & "</option>"
			oOrgs.movenext
		Loop 
		response.write "</select>"
	End If 

	oOrgs.close
	Set oOrgs = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowFeatures( iOrgId, iParentFeatureId )
'--------------------------------------------------------------------------------------------------
Sub ShowFeatures( iOrgId, iParentFeatureId )
	Dim sSql, oFeatures, iRows, sFeatureNotes, sFeatureDescription

	iRows = 0
	sSql = "Select featureid, featurename, haspublicview, publicurl, hasadminview, publicdisplayorder, publicimageurl, featuretype, isdefault, "
	sSql = sSql & " isnull(featurenotes,'') as featurenotes, isnull(featuredescription,'') as featuredescription from egov_organization_features where parentfeatureid = " & iParentFeatureId
	
	If iParentFeatureId = 0 Then 
		sSql = sSql & " Order By admindisplayorder"
	Else
		sSql = sSql & " Order By securitydisplayorder"
	End If 

	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 0, 1

	Do While Not oFeatures.EOF
		iRows = iRows + 1
		response.write vbcrlf & "<tr"
		If iParentFeatureId = 0 Then 
			response.write " class=""featureparent"" "
		Else 
			If iRows Mod 2 = 0 Then response.write " class=""altrow"" "
		End If 
		response.write ">"
		response.write "<td>"
		response.write vbcrlf & "<input type=""checkbox"" name=""featureid"" value=""" & oFeatures("featureid") & """ "
		response.write CheckOrgHasFeature( iOrgId, oFeatures("featureid") )
'		If UCase(oFeatures("featurename")) = "HOME" Or UCase(oFeatures("featurename")) = "SECURITY" Then response.write " disabled=""disabled"" "
		response.write " /></td><td nowrap=""nowrap"">" 
		If iParentFeatureId > 0 Then response.write " &nbsp; &bull; "
		response.write oFeatures("featurename") & " &nbsp; <a href=""new_feature.asp?orgid=" & iOrgId & "&iFeatureId=" & oFeatures("featureid") & """><img src=""../images/edit.gif"" border=""0""></a></td>"
		response.write "<td align=""center"">" & oFeatures("featureid") & "</td>"
		response.write "<td align=""center"">" & oFeatures("featuretype") & "</td>"
		response.write "<td align=""center"">" & oFeatures("haspublicview") & "</td>"
		response.write "<td align=""center"">" & oFeatures("isdefault") & "</td>"
		response.write "<td>" 
		sFeatureNotes = ""
		sFeatureNotes = CStr(oFeatures("featurenotes"))
		If Trim(sFeatureNotes) = "" Then 
			response.write "&nbsp;"
		Else
			response.write sFeatureNotes
		End If 
		response.write "</td>"

		sFeatureDescription = ""
		sFeatureDescription = CStr(oFeatures("featuredescription"))
		GetOrgFeaturesOverrides iOrgId, oFeatures("featureid"), oFeatures("publicurl"), oFeatures("haspublicview"), oFeatures("publicdisplayorder"), oFeatures("publicimageurl"), sFeatureDescription
		response.write "</tr>"
		ShowFeatures iOrgId, oFeatures("featureid")
		oFeatures.movenext 
	Loop 

	oFeatures.close
	Set oFeatures = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetOrgFeaturesOverrides( iOrgId )
'--------------------------------------------------------------------------------------------------
Sub GetOrgFeaturesOverrides( iOrgId, iFeatureId, sDefaultPublicUrl, bHaspublicview, iDefaultpublicdisplayorder, sDefaultpublicimageurl, sDefaultfeaturedescription )
	Dim sSql, oOrgFeature, sDisplayDesc

	'If IsNull(sDefaultfeaturedescription) Then sDefaultfeaturedescription = ""
		
	sSql = "Select isnull(featurename,'NULL') as featurename, isnull(featuredescription,'NULL') as featuredescription, isnull(publicurl,'NULL') as publicurl, publicdisplayorder, isnull(publicimageurl,'NULL') as publicimageurl, "
	sSql = sSql & " hasadminview, admindisplayorder, publiccanview from egov_organizations_to_features where orgid = " 
	sSql = sSql & iOrgId & " and featureid = " & iFeatureId

	Set oOrgFeature = Server.CreateObject("ADODB.Recordset")
	oOrgFeature.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

	If Not oOrgFeature.EOF Then
		response.write "<td align=""center""><input type=""checkbox"" name=""publiccanview" & iFeatureId & """" 
		If oOrgFeature("publiccanview") Then 
			response.write " checked=""checked"" "
		End If 
		response.write " /></td>"
		response.write "<td><input type=""text"" name=""featurename" & iFeatureId & """ value=""" & oOrgFeature("featurename") & """ maxlength=""255"" /></td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then 
			response.write "<textarea cols=""30"" rows=""3""name=""featuredescription" & iFeatureId & """>" & oOrgFeature("featuredescription") & "</textarea>"
			'sDisplayDesc = Replace(sDefaultfeaturedescription,"'","&rsquo;")
			sDisplayDesc = Replace(sDefaultfeaturedescription,"'","\'")
			response.write "<br /><span class=""spanlink"" onclick=""alert('Default Feature Description:\r\n" & sDisplayDesc & "');"">[" & Left(sDefaultfeaturedescription,10) & "...]</span>"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicurl" & iFeatureId & """ value=""" & oOrgFeature("publicurl") & """ maxlength=""255"" />"
			response.write "<br />[" & sDefaultPublicUrl & "]"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicdisplayorder" & iFeatureId & """ value=""" 
			If IsNull(oOrgFeature("publicdisplayorder")) Then 
				response.write "NULL" 
			Else 
				response.write oOrgFeature("publicdisplayorder") 
			End If 
			response.write """ size=""4"" maxlength=""4"" />"
			response.write "<br />[" & iDefaultpublicdisplayorder & "]"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicimageurl" & iFeatureId & """ value=""" & oOrgFeature("publicimageurl") & """ maxlength=""255"" />"
			response.write "<br />[" & sDefaultpublicimageurl & "]"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
	Else 
		response.write "<td align=""center""><input type=""checkbox"" name=""publiccanview" & iFeatureId & """ /></td>"
		response.write "<td><input type=""text"" name=""featurename" & iFeatureId & """ value=""NULL"" maxlength=""255"" /></td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then
			response.write "<textarea cols=""30"" rows=""3""name=""featuredescription" & iFeatureId & """>NULL</textarea>"
			sDisplayDesc = Replace(sDefaultfeaturedescription,"'","\'")
			response.write "<br /><span class=""spanlink"" onclick=""alert('Default Feature Description:\r\n" &  sDisplayDesc & "');"">[" & Left(CStr(sDefaultfeaturedescription),10) & "...]</span>"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicurl" & iFeatureId & """ value=""NULL"" maxlength=""255"" />"
			response.write "<br />[" & sDefaultPublicUrl & "]"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicdisplayorder" & iFeatureId & """ value=""NULL"" size=""4"" maxlength=""4"" />"
			response.write "<br />[" & iDefaultpublicdisplayorder & "]"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicimageurl" & iFeatureId & """ value=""NULL"" maxlength=""255"" />"
			response.write "<br />[" & sDefaultpublicimageurl & "]"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
	End If 

	oOrgFeature.close
	Set oOrgFeature = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function CheckOrgHasFeature( iOrgId, iFeatureId )
'--------------------------------------------------------------------------------------------------
Function CheckOrgHasFeature( iOrgId, iFeatureId )
	Dim sSql, oOrgFeature

	sSql = "Select featureid from egov_organizations_to_features where orgid = " & iOrgId & " and featureid = " & iFeatureId

	Set oOrgFeature = Server.CreateObject("ADODB.Recordset")
	oOrgFeature.Open sSQL, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If Not oOrgFeature.EOF Then
		CheckOrgHasFeature = " checked=""checked"" "
	Else 
		CheckOrgHasFeature = ""
	End If 

	oOrgFeature.close
	Set oOrgFeature = Nothing

End Function 


%>


