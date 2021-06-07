<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: new_feature.asp
' AUTHOR: Steve Loar
' CREATED: 12/13/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is where features are added
'
' MODIFICATION HISTORY
' 1.0   12/13/06   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iFeatureId, sFeature, sFeatureName, sFeatureNotes, sFeatureDescription, sHasPublicView, sPublicUrl
Dim sPublicImageurl, sHasAdminView, sAdminPageUrl, sFeatureType, sHasPermissions, sHhasPermissionLevels
Dim sRootAdminRequired, sParentFeatureId, sIsDefault

sLevel = "../" ' Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then
	response.redirect "../default.asp"
End If 

If request("iFeatureId") = "" Or clng(request("iFeatureId")) = clng(0) Then
	iFeatureId = 0
	sFeatureType = "Q"  ' There should not be a 'Q'
	sParentFeatureId = 0
Else
	iFeatureId = clng(request("iFeatureId"))
	GetFeatureValues iFeatureId
End If 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />

	<script language="Javascript" src="tablesort.js"></script>

<script language="Javascript">
<!--

	function Validate() 
	{
		if (document.frmFeature.feature.value == '')
		{
			alert("Feature Name cannot be blank.");
			document.frmFeature.feature.focus();
			return;
		}
		if (document.frmFeature.featurename.value == '')
		{
			alert("Displayed Feature Name cannot be blank.");
			document.frmFeature.featurename.focus();
			return;
		}
		document.frmFeature.submit();
	}


//-->
</script>

</head>

<body>

 
<%'DrawTabs tabRecreation,1%>
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
<!--BEGIN: PAGE TITLE-->
<p>
	<font size="+1"><strong>
<%		If iFeatureId = 0 Then %>
			New 
<%		End If %>
		Organization Feature</strong></font><br />
</p>
<!--END: PAGE TITLE-->


<!--BEGIN: FUNCTION LINKS-->
<div id="functionlinks">
		<a href="manage_features.asp?orgid=<%=request("orgid")%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to Feature Management</a>&nbsp;&nbsp;
		<a href="javascript:Validate();"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;
<%		If iFeatureId = 0 Then %>
			Create 
<%		Else %>
			Update 
<%		End If %>
		Feature</a>&nbsp;&nbsp;
</div>
<!--END: FUNCTION LINKS-->


<!--BEGIN: EDIT FORM-->
<form name="frmFeature" action="new_feature_save.asp" method="post">
	<input type="hidden" name="orgid" value="<%=request("orgid")%>" />
	<input type="hidden" name="iFeatureId" value="<%=iFeatureId%>" />
<div class="shadow">
	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin">
		<tr>
			<th>Feature</th>
		</tr>
		<tr>
			<td>
				<table>
					<tr>
						<td align="right">Feature Name:<br />(Used in the code)</td>
<%						If iFeatureId = 0 Then %>
							<td><input type="text" name="feature" value="" size="25" maxlength="50" /></td>
<%						Else %>
							<td><input type="hidden" name="feature" value="<%=sFeature%>" /><%=sFeature%></td>
<%						End If %>
					</tr>
					<tr>
						<td align="right">Default Displayed Name:</td><td><input type="text" name="featurename" value="<%=sFeatureName%>" size="90" maxlength="255" /></td>
					</tr>
					<tr>
						<td align="right">Feature Notes:</td><td><textarea name="featurenotes" class="features"><%=sFeatureNotes%></textarea></td>
					</tr>
					<tr>
						<td align="right">Default Public Description:<br />(Displayed on Home Page)</td><td><textarea name="featuredescription" class="features"><%=sFeatureDescription%></textarea></td>
					</tr>
					<tr>
						<td align="right">&nbsp;</td><td><input type="checkbox" name="haspublicview" <%=sHasPublicView%> /> &nbsp; This is a Public Feature</td>
					</tr>
					<tr>
						<td align="right">Default Public Page:</td><td><input type="text" name="publicurl" value="<%=sPublicUrl%>" size="90" maxlength="512" /></td>
					</tr>
					<tr>
						<td align="right">Default Public Image:</td><td><input type="text" name="publicimageurl" value="<%=sPublicImageurl%>" size="90" maxlength="255" /></td>
					</tr>
					<tr>
						<td align="right"> &nbsp;</td><td><input type="checkbox" name="hasadminview" <%=sHasAdminView%> /> This is an Admin Feature</td>
					</tr>
					<tr>
						<td align="right">Admin Page URL:</td><td><input type="text" name="adminurl" value="<%=sAdminPageUrl%>" size="90" maxlength="255" /></td>
					</tr>
					<tr>
						<td align="right">Parent Feature:</td><td><% ShowParentFeatures %></td>
					</tr>
					<tr>
						<td align="right">Feature Type:</td>
						<td>
							<select name="featuretype">
								<option value="N"<% If sFeatureType = "N" Then response.write " selected=""selected"" " End If %>>Navigation</option>
								<option value="S"<% If sFeatureType = "S" Then response.write " selected=""selected"" " End If %>>Security<option>
							</select>
						</td>
					</tr>
					<tr>
						<td align="right"> &nbsp;</td><td><input type="checkbox" name="haspermissions" <%=sHasPermissions%> /> &nbsp; Users Need Permission Assigned</td>
					</tr>
					<tr>
						<td align="right"> &nbsp;</td><td><input type="checkbox" name="haspermissionlevels" <%=sHhasPermissionLevels%> /> &nbsp; This requires Permission Levels</td>
					</tr>
					<tr>
						<td align="right"> &nbsp;</td><td><input type="checkbox" name="rootadminrequired" <%=sRootAdminRequired%> /> &nbsp; Root Admin Status Required<br />(Use to restrict a feature to the root admin OR so that only the root admin can assign permissions.)</td>
					</tr>
					<tr>
						<td align="right"> &nbsp;</td><td><input type="checkbox" name="isdefault" <%=sIsDefault%> /> &nbsp; This is a Default Setup Feature</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>

</form>
<!--END: EDIT FORM-->

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
' Sub ShowParentFeatures()
'--------------------------------------------------------------------------------------------------
Sub ShowParentFeatures()
	Dim sSql, oFeatures

	sSql = "Select featureid, featurename from egov_organization_features where parentfeatureid = 0 Order by admindisplayorder"

	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 0, 1

	If Not oFeatures.EOF Then 
		response.write vbcrlf & "<select name=""parentfeatureid"">"
		response.write vbcrlf & " <option value=""0"">This is a top level feature</option>"
		Do While Not oFeatures.EOF
			response.write vbcrlf & " <option value=""" & oFeatures("featureid") & """"
			If clng(sParentFeatureId) = clng(oFeatures("featureid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oFeatures("featurename") & "</option>"
			oFeatures.MoveNext
		Loop 
		response.write vbcrlf & "</select>" & vbcrlf
	End If 

	oFeatures.close
	Set oFeatures = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub GetFeatureValues( iFeatureId )
'--------------------------------------------------------------------------------------------------
Sub GetFeatureValues( iFeatureId )
	Dim sSql, oFeatures

	sSql = "Select * from egov_organization_features where featureid = " & iFeatureId

	Set oFeatures = Server.CreateObject("ADODB.Recordset")
	oFeatures.Open sSQL, Application("DSN"), 0, 1
	
	If Not oFeatures.EOF Then 
		oFeatures.MoveFirst
		'response.write vbcrlf & "feature = [" & oFeatures("feature") & "]"
		sFeature = oFeatures("feature")
		'response.write vbcrlf & "<br />featurename = [" & oFeatures("featurename") & "]"
		sFeatureName = oFeatures("featurename")
		'response.write vbcrlf & "<br />featuredescription = [" & oFeatures("featuredescription") & "]"
		sFeatureDescription = oFeatures("featuredescription")
		
		'response.write vbcrlf & "<br />haspublicview = [" & oFeatures("haspublicview") & "]"
		If oFeatures("haspublicview") Then 
			sHasPublicView = " checked=""checked"" "
		Else
			sHasPublicView = ""
		End If 
		'response.write vbcrlf & "<br />publicurl = [" & oFeatures("publicurl") & "]"
		sPublicUrl = oFeatures("publicurl")
		'response.write vbcrlf & "<br />publicimageurl = [" & oFeatures("publicimageurl") & "]"
		sPublicImageurl = oFeatures("publicimageurl")
		'response.write vbcrlf & "<br />hasadminview = [" & oFeatures("hasadminview") & "]"
		
		If oFeatures("hasadminview") Then 
			sHasAdminView = " checked=""checked"" "
		Else
			sHasAdminView = ""
		End If 
		'response.write vbcrlf & "<br />adminurl = [" & oFeatures("adminurl") & "]"
		sAdminPageUrl = oFeatures("adminurl")
		'response.write vbcrlf & "<br />parentfeatureid = [" & oFeatures("parentfeatureid") & "]"
		sParentFeatureId = oFeatures("parentfeatureid")
		'response.write vbcrlf & "<br />featuretype = [" & oFeatures("featuretype") & "]"
		sFeatureType = oFeatures("featuretype")
		'response.write vbcrlf & "<br />haspermissions = [" & oFeatures("haspermissions") & "]"
		If oFeatures("haspermissions") Then 
			sHasPermissions = " checked=""checked"" "
		Else
			sHasPermissions = ""
		End If 
		'response.write vbcrlf & "<br />haspermissionlevels = [" & oFeatures("haspermissionlevels") & "]"
		If oFeatures("haspermissionlevels") Then 
			sHhasPermissionLevels = " checked=""checked"" "
		Else
			sHhasPermissionLevels = ""
		End If 
		'response.write vbcrlf & "<br />rootadminrequired = [" & oFeatures("rootadminrequired") & "]"
		If oFeatures("rootadminrequired") Then 
			sRootAdminRequired = " checked=""checked"" "
		Else
			sRootAdminRequired = ""
		End If 
		If oFeatures("isdefault") Then 
			sIsDefault = " checked=""checked"" "
		Else
			sIsDefault = ""
		End If 
		
		sFeatureNotes = ""
		'response.write vbcrlf & "<br />featurenotes = [" & oFeatures("featurenotes") & "]"
		sFeatureNotes = oFeatures("featurenotes")
	End If 

	oFeatures.close
	Set oFeatures = Nothing 
	'response.End 
End Sub 



%>


