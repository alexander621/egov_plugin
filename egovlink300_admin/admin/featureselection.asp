<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: featureselection.asp
' AUTHOR: Steve Loar
' CREATED: 08/25/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the management of features for clients
'
' MODIFICATION HISTORY
' 1.1	10/20/2011	Steve Loar - Converted to HTML5
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iMaxElement, sLoadMsg

sLevel = "../"  'Override of value from common.asp

If Not UserIsRootAdmin( session("UserID") ) Then 
  	response.redirect "../default.asp"
End If 

'If request("orgid") <> "" Then 
  	'iOrgId = request("orgid")
'Else 
   'If session("orgid") <> "" Then 
      iOrgID = session("orgid")
   'Else 
     	'iOrgId = GetMaxOrgId
   'End If 
'End If 

iMaxElement = GetMaxParentElementId()

If request("s") <> "" Then
	If request("s") = "upd" Then
		sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
	End If
End If 

%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-GovLink Administration Console {Organization Feature Selection}</title>
	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../classes/classes.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	<!--

		function displayScreenMsg( iMsg ) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "";
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>
		}

		function ShowOrg()
		{
			document.pickForm.submit();
		}

		function AdminSecurity( sOrgId )
		{
			if (confirm('Run Initial Security Setup?'))
			{
				window.location='admin_security.asp?orgid=' + sOrgId;
			}
		}

		function ChangeSelection( iOrgid, iFeatureId )
		{
			//alert("Org: " + iOrgid + "  Feature: " + iFeatureId );
			doAjax('featureselectionupdate.asp', 'orgid=' + iOrgid + '&featureid=' + iFeatureId, '', 'get', '0');
			//window.location='featureselectionupdate.asp?orgid=' + iOrgid + '&featureid=' + iFeatureId;
		}

		function togglethis( toggleid )
		{
			//alert($("toggle"+toggleid).innerHTML);

			// Change the plus and minus sign
			if($("toggle"+toggleid).innerHTML == '+')
			{
				$("toggle"+toggleid).innerHTML = '&ndash;';
			}
			else
			{
				$("toggle"+toggleid).innerHTML = '+';
			}

			
			// Show or hide the rows
			var elements = document.getElementsByClassName(toggleid);

			for (var i = 0; i < elements.length; i++ )
			{
				if (elements[i].style.display == '')
				{
					elements[i].style.display = 'none';
				}
				else
				{
					elements[i].style.display = '';
				}
			}
		}

		function ShowAll()
		{
			var iMaxElement = <%=iMaxElement%>;

			for (var x = 1; x < iMaxElement; x++ )
			{
				if ($("toggle"+x))
				{
					// Throw the + to a -
					$("toggle"+x).innerHTML = '&ndash;';

					// Get the child rows
					var elements = document.getElementsByClassName(x);

					// Show the child rows
					for (var i = 0; i < elements.length; i++ )
					{
						elements[i].style.display = '';
					}
				}
			}
		}

		function HideAll()
		{
			var iMaxElement = <%=iMaxElement%>;

			for (var x = 1; x < iMaxElement; x++ )
			{
				if ($("toggle"+x))
				{
					// Throw the - to a +
					$("toggle"+x).innerHTML = '+';

					// Get the child rows
					var elements = document.getElementsByClassName(x);

					// Show the child rows
					for (var i = 0; i < elements.length; i++ )
					{
						elements[i].style.display = 'none';
					}
				}
			}
		}

		document.getElementsByClassName = function(clsName)
		{    
			var retVal = new Array();    
			// for this I just want the table rows
			var elements = document.getElementsByTagName("tr");    
			for(var i = 0;i < elements.length;i++)
			{        
				if(elements[i].className.indexOf(" ") >= 0)
				{            
					var classes = elements[i].className.split(" ");            
					for(var j = 0;j < classes.length;j++)
					{                
						if(classes[j] == clsName)                    
							retVal.push(elements[i]);            
					}        
				}
				else if(elements[i].className == clsName)            
					retVal.push(elements[i]);    
			}    
			return retVal;
		}

	//-->
	</script>

</head>
<body onload="SetUpPage();">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

		<h3>Organization Feature Selection</h3>
		
		<!--form name="pickForm" method="post" action="featureselection.asp">
			<p>
				Organization: <% 'ShowOrgDropDown iOrgId %> &nbsp; <span id="screenMsg"></span><br /><br />
			</p>
		</form-->

		<form name="FeatureForm" method="post" action="save_features.asp">
			<input type="hidden" name="orgid" value="<%=iOrgId%>" />

			<div id="featurebuttons">
				<table border="0" cellspacing="0" cellpadding="2" id="featureselectiontopbuttons">
				<tr valign="top">
					<td><input type="submit" class="button" name="save1" value="Save Changes" /> &nbsp;&nbsp; &nbsp;&nbsp;</td>
					<td>
					<input type="button" class="button" name="manage" value="Feature Management" onclick="javascript:window.location='managefeatures.asp';" />
					<input type="button" class="button" name="edit" value="Org Properties" onClick="javascript:window.location='edit_org.asp?orgid=<%=iOrgId%>'" />
					<input type="button" class="button" name="createorg" value="New Org" onClick="javascript:window.location='new_org.asp?orgid=<%=iOrgId%>'" />
					<input type="button" class="button" name="setupdisplays" value="Setup Displays" onClick="javascript:window.location='clientdisplaylist.asp'" />
					<input type="button" class="button" name="docsync" value="Document Sync" onClick="javascript:window.location='../docs/docsync.asp?orgid=<%=iOrgId%>'" />
					<input type="button" class="button" name="adminsecurity" value="Set Initial Admin Security" onClick="AdminSecurity('<%=iOrgId%>')" />
					<br />
					<div id="lowerbuttons">
						<input type="button" class="button" name="showall" id="showall" value="Open All" onclick="ShowAll();" />
						<input type="button" class="button" name="hideall" id="hideall" value="Close All" onclick="HideAll();" />
						<input type="button" class="button" name="assignFeature" id="assignFeature" value="Assign Feature to Orgs" onClick="location.href='featureassign.asp?assigntype=ORG';" />
						<input type="button" class="button" name="assignFeature" id="assignFeature" value="Assign Feature to Users" onClick="location.href='featureassign.asp?assigntype=USER';" />
						<input type="button" class="button" name="mobilefeatures" value="Mobile Feature Setup" onClick="javascript:window.location='mobilefeaturesetup.asp?orgid=<%=iOrgId%>'" />
					</div>
					</td>
				</tr>
				</table>
			</div>

			<table id="featuremanager" border="0" cellpadding="3" cellspacing="0">
				<tr><th colspan="2">Feature</th><th>Id</th><th>Feature<br />Type</th><th>Public Feature</th><th>Default Feature</th>
					<th width="250" nowrap="nowrap">Feature Notes</th><th>Public Can View</th><th>Feature Name</th><th>Feature Description</th>
					<th>Public URL</th><th>Public Display Order</th><th>Public Image URL</th>
				</tr>
<% 
				ShowFeatures iOrgId, 0 
%>
			</table>

			<div id="bottomsave">
				<p>
					<input type="submit" class="button" value="Save Changes" />
				</p>
			</div>
		</form>
		
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' integer GetMaxOrgId()
'------------------------------------------------------------------------------
Function GetMaxOrgId()
	Dim sSql, oRs

	sSql = "SELECT MAX(orgid) AS maxorgid FROM organizations"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	GetMaxOrgId = oRs("maxorgid")

	oRs.Close
	Set oRs = Nothing

End Function 

'------------------------------------------------------------------------------
' void ShowOrgDropDown iOrgId 
'------------------------------------------------------------------------------
Sub ShowOrgDropDown( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT orgname, orgcity, orgid, defaultstate FROM organizations "
	sSql = sSql & "WHERE isdeactivated = 0 ORDER BY orgcity, defaultstate"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), adOpenForwardOnly, adLockReadOnly

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""orgid"" onchange='ShowOrg();'>"
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("orgid") & """ "
			If CLng(iOrgId) = CLng(oRs("orgid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("orgcity") & ", " & oRs("defaultstate") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ShowFeatures iOrgId, iParentFeatureId 
'------------------------------------------------------------------------------
Sub ShowFeatures( ByVal iOrgId, ByVal iParentFeatureId )
	Dim sSql, oRs, iRows, sFeatureNotes, sFeatureDescription

	iRows = 0
	sSql = "SELECT featureid, featurename, haspublicview, publicurl, hasadminview, publicdisplayorder, publicimageurl, featuretype, "
	sSql = sSql & " isdefault, ISNULL(featurenotes,'') AS featurenotes, ISNULL(featuredescription,'') AS featuredescription, "
	sSql = sSql & " hasmobileview, ISNULL(mobileiframeurl,'') AS mobileiframeurl, mobiledefaultitemcount, "
	sSql = sSql & " ISNULL(mobileurl,'') AS mobileurl, mobiledefaultdisplayorder, mobiledefaultlistcount, ismobilenavonly "
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

		checkForCLInfo oRs("featureid"), iorgid, lcl_cl_postcomments_label, lcl_cl_postcomments_formid

		response.write vbcrlf & "<tr"

		If iParentFeatureId = 0 Then 
			response.write " class=""featureparent"" "
		Else 
			response.write " class=""" & iParentFeatureId
			If iRows Mod 2 = 0 Then 
				response.write " altrow"
			End If 
			response.write """ style=""display: none;"" "
		End If 

		response.write ">"
		response.write "<td valign=""top"">"
		response.write vbcrlf & "<input type=""checkbox"" name=""featureid"" value=""" & oRs("featureid") & """ "
		response.write CheckOrgHasFeature( iOrgId, oRs("featureid") )
		response.write " onclick=""ChangeSelection( " & iOrgId & ", " & oRs("featureid") & " );"""
		response.write " />" & vbcrlf

		response.write "<input type=""hidden"" name=""CL_postcomments_label"  & oRs("featureid") & """ id=""CL_postcomments_label"  & oRs("featureid") & """ value=""" & lcl_cl_postcomments_label  & """ />" & vbcrlf
		response.write "<input type=""hidden"" name=""CL_postcomments_formid" & oRs("featureid") & """ id=""CL_postcomments_formid" & oRs("featureid") & """ value=""" & lcl_cl_postcomments_formid & """ />" & vbcrlf

		response.write "</td>"

		 
		If iParentFeatureId > 0 Then 
			response.write "<td nowrap=""nowrap"" valign=""top"">"
			response.write " &nbsp; &bull; " & oRs("featurename")
		Else
			response.write "<td nowrap=""nowrap"" valign=""top"" onclick=""togglethis('" & oRs("featureid") & "');"">"
			response.write "<span id=""toggle" & oRs("featureid") & """ >+</span>&nbsp;"
			response.write "<strong>" & oRs("featurename") & " </strong>"
		End If 
		response.write "</td>"

		response.write "<td valign=""top"" align=""center"">" & oRs("featureid") & "</td>"

		response.write "<td align=""center"" valign=""top"">"
		If oRs("featuretype") = "N" Then 
			response.write "Navigation"
		Else
			response.write "Security"
		End If 
		response.write "</td>"

		response.write "<td align=""center"" valign=""top"">"
		If oRs("haspublicview") Then 
			response.write "YES"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "<td align=""center"" valign=""top"">"
		If oRs("isdefault") Then 
			response.write "YES"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "<td valign=""top"">" 
		sFeatureNotes = ""
		sFeatureNotes = CStr(oRs("featurenotes"))
		If Trim(sFeatureNotes) = "" Then 
			response.write "&nbsp;"
		Else
			response.write sFeatureNotes
		End If 
		response.write "</td>"

		sFeatureDescription = ""
		sFeatureDescription = CStr(oRs("featuredescription"))
		GetOrgFeaturesOverrides iOrgId, oRs("featureid"), oRs("publicurl"), oRs("haspublicview"), oRs("publicdisplayorder"), oRs("publicimageurl"), sFeatureDescription

		response.write "</tr>"

		' Get any sub features
		ShowFeatures iOrgId, oRs("featureid")

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void GetOrgFeaturesOverrides iOrgId, iFeatureId, sDefaultPublicUrl, bHaspublicview, iDefaultpublicdisplayorder, sDefaultpublicimageurl, sDefaultfeaturedescription
'------------------------------------------------------------------------------
Sub GetOrgFeaturesOverrides( ByVal iOrgId, ByVal iFeatureId, ByVal sDefaultPublicUrl, ByVal bHaspublicview, ByVal iDefaultpublicdisplayorder, ByVal sDefaultpublicimageurl, ByVal sDefaultfeaturedescription )
	Dim sSql, oRs, sDisplayDesc, iMobileIsActivated

	'If IsNull(sDefaultfeaturedescription) Then sDefaultfeaturedescription = ""
		
	sSql = "SELECT ISNULL(featurename,'NULL') AS featurename, ISNULL(featuredescription,'NULL') AS featuredescription, "
	sSql = sSql & " ISNULL(publicurl,'NULL') AS publicurl, publicdisplayorder, ISNULL(publicimageurl,'NULL') AS publicimageurl, "
	sSql = sSql & " hasadminview, admindisplayorder, publiccanview, mobileisactivated, ISNULL(mobilename,'') AS mobilename, "
	sSql = sSql & " ISNULL(mobiledisplayorder,0) AS mobiledisplayorder, ISNULL(mobileitemcount,0) AS mobileitemcount, ISNULL(mobilelistcount,0) AS mobilelistcount "
	sSql = sSql & " FROM egov_organizations_to_features "
	sSql = sSql & " WHERE orgid = " & iOrgId & " AND featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write "<td align=""center"" valign=""top""><input type=""checkbox"" name=""publiccanview" & iFeatureId & """" 
		If oRs("publiccanview") Then 
			response.write " checked=""checked"" "
		End If 
		response.write " />"
		If oRs("mobileisactivated") Then
			iMobileIsActivated = 1
		Else
			iMobileIsActivated = 0
		End If 
		response.write "<input type=""hidden"" name=""mobileisactivated" & iFeatureId & """ value=""" & iMobileIsActivated & """ />"
		response.write "<input type=""hidden"" name=""mobilename" & iFeatureId & """ value=""" & oRs("mobilename") & """ />"
		response.write "<input type=""hidden"" name=""mobiledisplayorder" & iFeatureId & """ value=""" & oRs("mobiledisplayorder") & """ />"
		response.write "<input type=""hidden"" name=""mobileitemcount" & iFeatureId & """ value=""" & oRs("mobileitemcount") & """ />"
		response.write "<input type=""hidden"" name=""mobilelistcount" & iFeatureId & """ value=""" & oRs("mobilelistcount") & """ />"
		response.write "</td>"

		response.write "<td valign=""top""><input type=""text"" name=""featurename" & iFeatureId & """ value=""" & oRs("featurename") & """ maxlength=""255"" /></td>"

		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then 
			response.write "<textarea cols=""30"" rows=""3""name=""featuredescription" & iFeatureId & """>" & oRs("featuredescription") & "</textarea>"
			'sDisplayDesc = Replace(sDefaultfeaturedescription,"'","&rsquo;")
			sDisplayDesc = Replace(sDefaultfeaturedescription,"'","\'")
			response.write "<br /><span class=""spanlink"" onclick=""alert('Default Feature Description:\r\n" & sDisplayDesc & "');"">[" & Left(sDefaultfeaturedescription,10) & "...]</span>"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicurl" & iFeatureId & """ value=""" & oRs("publicurl") & """ maxlength=""255"" />"
			response.write "<br />[" & sDefaultPublicUrl & "]"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicdisplayorder" & iFeatureId & """ value=""" 
			If IsNull(oRs("publicdisplayorder")) Then 
				response.write "NULL" 
			Else 
				response.write oRs("publicdisplayorder") 
			End If 
			response.write """ size=""4"" maxlength=""4"" />"
			response.write "<br />[" & iDefaultpublicdisplayorder & "]"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicimageurl" & iFeatureId & """ value=""" & oRs("publicimageurl") & """ maxlength=""255"" />"
			response.write "<br />[" & sDefaultpublicimageurl & "]"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
	Else 
		response.write "<td align=""center"" valign=""top""><input type=""checkbox"" name=""publiccanview" & iFeatureId & """ />"
		response.write "<input type=""hidden"" name=""mobileisactivated" & iFeatureId & """ value=""0"" />"
		response.write "<input type=""hidden"" name=""mobilename" & iFeatureId & """ value="""" />"
		response.write "<input type=""hidden"" name=""mobiledisplayorder" & iFeatureId & """ value=""0"" />"
		response.write "<input type=""hidden"" name=""mobileitemcount" & iFeatureId & """ value=""0"" />"
		response.write "<input type=""hidden"" name=""mobilelistcount" & iFeatureId & """ value=""0"" />"
		response.write "</td>"
		response.write "<td valign=""top""><input type=""text"" name=""featurename" & iFeatureId & """ value=""NULL"" maxlength=""255"" /></td>"
		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then
			response.write "<textarea cols=""30"" rows=""3""name=""featuredescription" & iFeatureId & """>NULL</textarea>"
			sDisplayDesc = Replace(sDefaultfeaturedescription,"'","\'")
			response.write "<br /><span class=""spanlink"" onclick=""alert('Default Feature Description:\r\n" &  sDisplayDesc & "');"">[" & Left(CStr(sDefaultfeaturedescription),10) & "...]</span>"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicurl" & iFeatureId & """ value=""NULL"" maxlength=""255"" />"
			response.write "<br />[" & sDefaultPublicUrl & "]"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicdisplayorder" & iFeatureId & """ value=""NULL"" size=""4"" maxlength=""4"" />"
			response.write "<br />[" & iDefaultpublicdisplayorder & "]"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td align=""center"" valign=""top"">"
		If bHaspublicview Then 
			response.write "<input type=""text"" name=""publicimageurl" & iFeatureId & """ value=""NULL"" maxlength=""255"" />"
			response.write "<br />[" & sDefaultpublicimageurl & "]"
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</td>"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' string CheckOrgHasFeature( iOrgId, iFeatureId )
'------------------------------------------------------------------------------
Function CheckOrgHasFeature( ByVal iOrgId, ByVal iFeatureId )
	Dim sSql, oRs

	sSql = "SELECT featureid "
	sSql = sSql & " FROM egov_organizations_to_features "
	sSql = sSql & " WHERE orgid = "   & iOrgId
	sSql = sSql & " AND featureid = " & iFeatureId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		CheckOrgHasFeature = " checked=""checked"" "
	Else 
		CheckOrgHasFeature = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void checkForCLInfo iFeatureID, p_orgid, lcl_cl_postcomments_label, lcl_cl_postcomments_formid
'------------------------------------------------------------------------------
Sub checkForCLInfo( ByVal iFeatureID, ByVal p_orgid, ByRef lcl_cl_postcomments_label, ByRef lcl_cl_postcomments_formid )
	Dim sSql, oRs

	lcl_cl_postcomments_label  = ""
	lcl_cl_postcomments_formid = 0

	sSql = "SELECT ISNULL(cl_postcomments_label,'') AS cl_postcomments_label, "
	sSql = sSql & " ISNULL(cl_postcomments_formid, 0) AS cl_postcomments_formid "
	sSql = sSql & " FROM egov_organizations_to_features "
	sSql = sSql & " WHERE orgid = "   & p_orgid
	sSql = sSql & " AND featureid = " & iFeatureID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		lcl_cl_postcomments_label  = oRs("cl_postcomments_label")
		lcl_cl_postcomments_formid = oRs("cl_postcomments_formid")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

'------------------------------------------------------------------------------
' integer GetMaxParentElementId()
'------------------------------------------------------------------------------
Function GetMaxParentElementId()
	Dim sSql, oRs

	sSql = "SELECT MAX(featureid) AS maxfeatureid FROM egov_organization_features WHERE parentfeatureid = 0"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetMaxParentElementId = oRs("maxfeatureid")
	Else
		GetMaxParentElementId = 0
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function



%>
