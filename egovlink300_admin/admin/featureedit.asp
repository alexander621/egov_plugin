<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="feature_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: featureedit.asp
' AUTHOR: Steve Loar
' CREATED: 9/12/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is where features are edited
'
' MODIFICATION HISTORY
' 1.0	09/12/2008	Steve Loar - Initial Version
' 1.1	04/17/2009	David Boyer - Added "Show on Community Link" options
' 1.2	07/20/2009	Steve Loar - Fixed to work in FireFox added Prototype. Javascript bug and formatting issues
' 1.3	07/27/2009	David Boyer - Added screen messages
' 1.4	07/29/2009	David Boyer - Added "Assign Feature to Orgs" button
' 2.0	04/05/2011	Steve Loar - Added Mobile Properties
' 2.1	10/20/2011	Steve Loar - Rearrange the page items a little.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iFeatureId, sFeature, sFeatureName, sFeatureNotes, sFeatureDescription, sHasPublicView, sPublicUrl
Dim sPublicImageurl, sHasAdminView, sAdminPageUrl, sFeatureType, sHasPermissions, sHhasPermissionLevels
Dim sRootAdminRequired, sParentFeatureId, sIsDefault, sCommunityLinkOn, sCL_numListItems, sPortalType
Dim sHasMobileView, sMobileURL, sIsMobileNavOnly, sMobileDefaultItemCount, sMobileDefaultDisplayOrder
Dim sMobileDefaultListCount

sLevel = "../"  'Override of value from common.asp

sMobileDefaultItemCount = ""
sMobileDefaultListCount = ""
sMobileDefaultDisplayOrder = ""

If Not UserIsRootAdmin(session("userid")) Then 
response.redirect "../default.asp"
End If 

iFeatureId = CLng(request("featureid"))

If iFeatureId = CLng(0) Then 
sFeatureType = "Q"  'There should not be a 'Q'
sParentFeatureId = 0
Else 
GetFeatureValues iFeatureId
End If 

'Check for a screen message
lcl_onload = "enableDisableCLField();"
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

	<title>E-Gov Administration Console {Organization Features: Edit}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script language="javascript" src="tablesort.js"></script>
	<script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="javascript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript">
	<!--

		function Validate() {
			var lcl_false_count = 0;

			if ($F("featurename") == '') 
			{
				lcl_focus = $("featurename");
				inlineMsg($("featurename").id,'<strong>Required Field Missing:</strong> Displayed Feature Name.',8,'featurename');
				lcl_false_count = lcl_false_count + 1;
			}
			else
			{
				clearMsg('featurename');
			}

			if ($F("feature") == '') 
			{
				lcl_focus = $("feature");
				inlineMsg($("feature").id,'<strong>Required Field Missing:</strong> Feature Name.',8,'feature');
				lcl_false_count = lcl_false_count + 1;
			}
			else
			{
				clearMsg('feature');
			}

			if(lcl_false_count > 0) 
			{
				lcl_focus.focus();
				return false;
			}
			else
			{
				if ($("iFeatureId").value == "0")
				{
					// Do a check that the feature name is unique before saving it
					doAjax('checkfeatureisunique.asp', 'feature=' + $("feature").value, 'featureCheckReturn', 'get', '0');
				}
				else
				{
					$("frmFeature").submit();
					return true;
				}
			}
		} 

		function featureCheckReturn( sResults )
		{
			//alert( sResults );
			if (sResults == "YES")
			{
				$("frmFeature").submit();
				return true;
			}
			else
			{
				$("feature").focus();
				inlineMsg($("feature").id,'<strong>Duplicate Field Value:</strong> The Feature Name must be unique and this one is already in use. Please try another name.',8,'feature');
				return false;
			}
		}

		function enableDisableCLField() 
		{
			var lcl_isChecked = $("CommunityLinkOn").checked;

			if(lcl_isChecked) 
			{
				$("CL_numListItems").disabled = false;
				$("CL_portaltype").disabled   = false;
			}
			else
			{
				$("CL_numListItems").disabled = true;
				$("CL_portaltype").disabled   = true;
				if($F("CL_numListItems") == "") 
				{
					$("CL_numListItems").value = 0;
				}
			}
		}

		function assignFeature(iFeatureID,iAssignType) 
		{
			location.href="featureassign.asp?featureid=" + iFeatureID + "&assigntype=" + iAssignType;
		}

		function displayScreenMsg(iMsg) {
			if(iMsg!="") 
			{
				document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
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
<form name="frmFeature" id="frmFeature" action="featureupdate.asp" method="post">
 	<input type="hidden" name="iFeatureId" id="iFeatureId" value="<%=iFeatureId%>" />
<div id="content">
 	<div id="centercontent">
<%
	'BEGIN: Page Title -----------------------------------------------------------
	If iFeatureId = 0 Then 
		lcl_title = "New"
	Else 
		lcl_title = "Edit"
	End If 

	response.write "<p>"
	response.write "<font size=""+1""><strong>" & lcl_title & " Feature</strong></font><br />"
	response.write "<input type=""button"" name=""returnButton"" id=""returnButton"" value=""<< Return to Feature Management"" class=""button"" style=""margin-top:5px;"" onclick=""location.href='managefeatures.asp';"" />"
	response.write "</p>"
 'END: Page Title -------------------------------------------------------------

 'BEGIN: Edit Form ------------------------------------------------------------
%>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr>
		<td colspan="2" align="right">&nbsp;<span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
	</tr>
	<tr>
		<td colspan="2"><% displayButtons iFeatureID %></td>
	</tr>
</table>

<table id="featuretable" cellpadding="5" cellspacing="0" border="0" class="tableadmin">
  <tr>
      <th>Feature Properties</th>
  </tr>
  <tr>
      <td>
          <table cellpadding="2" cellspacing="0" border="0">
            <tr>
                <td align="right" nowrap="nowrap">Feature Name:<br />(Used in the code)</td>
                <td>
              <%
                If CLng(iFeatureID) = CLng(0) Then 
                   lcl_shownHidden_feature = "text"
                   lcl_value_feature       = ""
                Else 
                   lcl_shownHidden_feature = "hidden"
                   lcl_value_feature       = sFeature
                End If 

                response.write "<input type=""" & lcl_shownHidden_feature & """ name=""feature"" id=""feature"" value=""" & lcl_value_feature & """ size=""25"" maxlength=""50"" />" & lcl_value_feature 
              %>
                </td>
            </tr>
            <tr>
                <td align="right">Default Displayed Name:</td>
                <td><input type="text" id="featurename" name="featurename" value="<%=sFeatureName%>" size="90" maxlength="255" /></td>
            </tr>
            <tr>
                <td align="right">Feature Notes:</td>
                <td><textarea name="featurenotes" class="features"><%=sFeatureNotes%></textarea></td>
            </tr>
			<tr>
                <td align="right">Parent Feature:</td>
                <td><% ShowParentFeatures %></td>
            </tr>
            <tr>
                <td align="right">Feature Type:</td>
                <td>
                    <select name="featuretype">
                      <option value="N"<% If sFeatureType = "N" Then response.write " selected=""selected"" " End If %>>Navigation</option>
                      <option value="S"<% If sFeatureType = "S" Then response.write " selected=""selected"" " End If %>>Security</option>
                    </select>
                </td>
            </tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td><input type="checkbox" name="haspermissions"<%=sHasPermissions%> /> &nbsp; Users Need Permission Assigned</td>
            </tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td><input type="checkbox" name="haspermissionlevels"<%=sHhasPermissionLevels%> /> &nbsp; This requires Permission Levels</td>
            </tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td><input type="checkbox" name="rootadminrequired"<%=sRootAdminRequired%> /> &nbsp; Root Admin Status Required<br />(Use to restrict a feature to the root admin OR so that only the root admin can assign permissions.)</td>
            </tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td><input type="checkbox" name="isdefault"<%=sIsDefault%> /> &nbsp; This is a Default Setup Feature</td>
            </tr>
            
			<tr class="featuregrouptitlerow">
				<td align="right">&nbsp;</td>
                <td><strong>Public Properties</strong>
				</td>
			</tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td><input type="checkbox" name="haspublicview" <%=sHasPublicView%> /> &nbsp; This is a Public Feature</td>
            </tr>
			<tr>
                <td align="right">Default Public Description:<br />(Displayed on Home Page)</td>
                <td><textarea name="featuredescription" class="features"><%=sFeatureDescription%></textarea></td>
            </tr>
            <tr>
                <td align="right">Default Public Page:</td>
                <td><input type="text" name="publicurl" value="<%=sPublicUrl%>" size="90" maxlength="512" /></td>
            </tr>
            <tr>
                <td align="right">Default Public Image:</td>
                <td><input type="text" name="publicimageurl" value="<%=sPublicImageurl%>" size="90" maxlength="255" /></td>
            </tr>

			<tr class="featuregrouptitlerow">
				<td align="right">&nbsp;</td>
                <td><strong>Admin Properties</strong>
				</td>
			</tr>
            <tr>
                <td align="right"> &nbsp;</td>
                <td><input type="checkbox" name="hasadminview" <%=sHasAdminView%> /> This is an Admin Feature</td>
            </tr>
            <tr>
                <td align="right">Admin Page URL:</td>
                <td><input type="text" name="adminurl" value="<%=sAdminPageUrl%>" size="90" maxlength="255" /></td>
            </tr>
			<tr class="featuregrouptitlerow">
				<td align="right">&nbsp;</td>
                <td><strong>Community Link Properties</strong>
				</td>
			</tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td><input type="checkbox" name="CommunityLinkOn" id="CommunityLinkOn"<%=sCommunityLinkOn%> onclick="enableDisableCLField()" /> &nbsp; This feature can be displayed on Community Link</td>
            </tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td>
                    <input type="text" name="CL_numListItems" id="CL_numListItems" value="<%=sCL_numListItems%>" size="2" maxlength="10" style="float:left; margin-right:5px;"/>
                    The default number of of list items that will be displayed for this feature (portal section) on the Community Link 
                    screen unless/until customized by org.
                </td>
            </tr>
            <tr>
                <td align="right">&nbsp;</td>
                <td>
                    <input type="text" name="CL_portaltype" id="CL_portaltype" value="<%=trim(sPortalType)%>" size="20" maxlength="50" style="float:left; margin-right:5px;"/>
                    <i>(used by developers ONLY)</i>&nbsp;Unique ID used on Community Link so that system knows which feature to build/display.
                </td>
            </tr>

			<tr class="featuregrouptitlerow">
				<td align="right">&nbsp;</td>
                <td><strong>Mobile Properties</strong>
				</td>
			</tr>
			<tr>
				<td align="right">&nbsp;</td>
                <td>
					<input type="checkbox" id="hasmobileview" name="hasmobileview"<%=sHasMobileView%> /> &nbsp; This feature has mobile pages. Check to display them.
				</td>
			</tr>
			<tr>
				<td align="right">&nbsp;</td>
                <td>
					<input type="checkbox" id="ismobilenavonly" name="ismobilenavonly"<%=sIsMobileNavOnly%> /> &nbsp; This feature only appears in the navigation. 
					It does not show items on the main mobile page.
				</td>
			</tr>
			<tr>
				<td align="right" valign="top">Main Page Display Order:&nbsp;</td>
                <td>
					<input type="text" id="mobiledefaultdisplayorder" name="mobiledefaultdisplayorder" value="<%=sMobileDefaultDisplayOrder%>" size="4" maxlength="4" />&nbsp;
					This is the feature's display position on the main mobile page and in navigation.
				</td>
			</tr>
			<tr>
				<td align="right" valign="top">Main Page Items:&nbsp;</td>
                <td>
					<input type="text" id="mobiledefaultitemcount" name="mobiledefaultitemcount" value="<%=sMobileDefaultItemCount%>" size="4" maxlength="4" />&nbsp;
					This is the number of items to show on the main mobile page.
				</td>
			</tr>
			<tr>
				<td align="right" valign="top">Mobile URL:&nbsp;</td>
                <td>
					<input type="text" id="mobileurl" name="mobileurl" value="<%=sMobileURL%>" size="90" maxlength="500" /><br />
					(This is normally the default page for when the feature is clicked from the Navigation.)
				</td>
			</tr>
			<tr>
				<td align="right" valign="top">Mobile URL Items:&nbsp;</td>
                <td>
					<input type="text" id="mobiledefaultlistcount" name="mobiledefaultlistcount" value="<%=sMobileDefaultListCount%>" size="4" maxlength="4" />&nbsp;
					This is the number of items to show on the feature's mobile page.
				</td>
			</tr>
          </table>
      </td>
  </tr>
</table>


<% displayButtons iFeatureID %>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  
</form>
</body>
</html>
<%

'------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' void ShowParentFeatures
'------------------------------------------------------------------------------
Sub ShowParentFeatures()
	Dim sSql, oRs, lcl_selected_parent

	sSql = "SELECT featureid, featurename "
	sSql = sSql & " FROM egov_organization_features "
	sSql = sSql & " WHERE parentfeatureid = 0 "
	sSql = sSql & " ORDER BY admindisplayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""parentfeatureid"" id=""parentfeatureid"">" 
		response.write vbcrlf & "<option value=""0"">This is a top level feature</option>" 

		Do While Not oRs.EOF
			If clng(sParentFeatureId) = clng(oRs("featureid")) Then 
				lcl_selected_parent = " selected=""selected"" "
			Else 
				lcl_selected_parent = ""
			End If 
			response.write vbcrlf & "<option value=""" & oRs("featureid") & """" & lcl_selected_parent & ">" & oRs("featurename") & "</option>" 
			oRs.MoveNext
		Loop 

		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' void GetFeatureValues iFeatureId
'------------------------------------------------------------------------------
Sub GetFeatureValues( ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT * FROM egov_organization_features "
	sSql = sSql & " WHERE featureid = " & iFeatureId

'	response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sFeature = oRs("feature")
		sFeatureName = oRs("featurename")
		sFeatureDescription = oRs("featuredescription")

		If oRs("haspublicview") Then 
			sHasPublicView = " checked=""checked"""
		Else 
			sHasPublicView = ""
		End If 

		sPublicUrl      = oRs("publicurl")
		sPublicImageurl = oRs("publicimageurl")

		If oRs("hasadminview") Then 
			sHasAdminView = " checked=""checked"""
		Else 
			sHasAdminView = ""
		End If 

		sAdminPageUrl    = oRs("adminurl")
		sParentFeatureId = oRs("parentfeatureid")
		sFeatureType     = oRs("featuretype")

		If oRs("haspermissions") Then 
			sHasPermissions = " checked=""checked"""
		Else 
			sHasPermissions = ""
		End If 

		If oRs("haspermissionlevels") Then 
			sHhasPermissionLevels = " checked=""checked"""
		Else 
			sHhasPermissionLevels = ""
		End If 

		If oRs("rootadminrequired") Then 
			sRootAdminRequired = " checked=""checked"""
		Else 
			sRootAdminRequired = ""
		End If 

		If oRs("isdefault") Then 
			sIsDefault = " checked=""checked"""
		Else 
			sIsDefault = ""
		End If 

		sFeatureNotes = ""
		sFeatureNotes = oRs("featurenotes")

		If oRs("CommunityLinkOn") Then 
			sCommunityLinkOn = " checked=""checked"""
		Else 
			sCommunityLinkOn = ""
		End If 

		sCL_numListItems = oRs("CL_numListItems")
		sPortalType = trim(oRs("CL_portaltype"))

		' Populate Mobile Variables
		If oRs("hasmobileview") Then
			sHasMobileView = " checked=""checked"""
		Else
			sHasMobileView = ""
		End If 

		sMobileURL = oRs("mobileurl")

		If oRs("ismobilenavonly") Then
			sIsMobileNavOnly = " checked=""checked"""
		Else
			sIsMobileNavOnly = ""
		End If 

		sMobileDefaultItemCount = oRs("mobiledefaultitemcount")

		sMobileDefaultListCount = oRs("mobiledefaultlistcount")

		sMobileDefaultDisplayOrder = oRs("mobiledefaultdisplayorder")
		' End of Mobile Variables

	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void displayButtons iFeatureid 
'------------------------------------------------------------------------------
Sub displayButtons( ByVal iFeatureid )
	Dim sButtonLabel

	If CLng(iFeatureid) = CLng(0) Then 
		sButtonLabel = "Create Feature"
	Else 
		sButtonLabel = "Save Changes"
	End If 

	response.write "<div id=""functionlinks"">"

	If CLng(iFeatureid) > CLng(0) Then 
		response.write "<input type=""button"" class=""button"" name=""assignFeature"" id=""assignFeature"" value=""Assign Feature to Orgs"" onclick=""assignFeature('"  & iFeatureid & "','ORG');"" />"
		response.write "&nbsp;<input type=""button"" class=""button"" name=""assignFeature"" id=""assignFeature"" value=""Assign Feature to Users"" onclick=""assignFeature('" & iFeatureid & "','USER');"" />"
	End If 

	response.write "&nbsp;<input type=""button"" name=""saveCreateButton"" id=""saveCreateButton"" value=""" & sButtonLabel & """ class=""button"" onclick=""Validate();"" />"
	response.write "</div>"

End Sub


%>
