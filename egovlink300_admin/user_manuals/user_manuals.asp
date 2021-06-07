<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: user_manuals.asp
' AUTHOR: ???
' CREATED: ???
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is the maintain action line code sections
'
' MODIFICATION HISTORY
' 1.0  08/20/08  David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("user_manuals") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../" ' Override of value from common.asp

'Determine which manual to open
 lcl_manual     = request("ID")
 lcl_manual_url = ""

'Verify that the org and user have access to the manual
 if NOT orghasfeature("user_manuals") then
    response.redirect sLevel & "permissiondenied.asp"
 else
    if not UserHasPermission( session("userid"), lcl_manual ) then
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if
%>
<html>
<head>
  <title>E-Gov Link {User Manuals}</title>
	
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />	
	<link rel="stylesheet" type="text/css" href="../global.css" />

<script language="javascript">
  function openManual(p_manual) {
     if(p_manual=="egov_manual") {
        //lcl_url = "E-GovDocumentation06.2008.pdf"
        lcl_url = "E-GovDocumentation.pdf"
     }else if(p_manual=="quickstartguide") {
        lcl_url = "QuickStartGuide.pdf"
     }
     window.open(lcl_url);
  }
</script>
</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="openManual('<%=lcl_manual%>');">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
  <div id="centercontent">

<p><h3>User Manuals</h3></p>

<div class="shadow">
<table border="0" cellpadding="5" cellspacing="0" class="tablelist">
  <tr><th colspan="2">&nbsp;</th></tr>
  <% displayUserManuals %>
</table>
</div>

  </div>
</div>

<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
sub displayUserManuals
 'Retrieve all of the user manuals this org and user have access to.
  sSQL = "SELECT isnull(otf.featurename,o.featurename) as featurename, o.feature, o.featureid, o.adminurl "
  sSQL = sSQL & " FROM egov_organization_features o, egov_organizations_to_features otf "
  sSQL = sSQL & " WHERE o.featureid = otf.featureid "
  sSQL = sSQL & " AND otf.orgid = " & session("orgid")
  sSQL = sSQL & " AND o.parentfeatureid = (select of2.featureid "
  sSQL = sSQL &                           " from egov_organization_features of2 "
  sSQL = sSQL &                           " where of2.feature = 'user_manuals') "
  sSQL = sSQL & " ORDER BY o.securitydisplayorder "

 	set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 3, 1

  if not rs.eof then
     lcl_bgcolor = "#eeeeee"
     while not rs.eof
        response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "    <td>" & rs("featurename") & ":</td>" & vbcrlf
        response.write "    <td align=""right""><input type=""button"" value=""View Manual"" onclick=""openManual('" & rs("feature") & "')""></td>" & vbcrlf
        response.write "</tr>" & vbcrlf

        lcl_bgcolor = changeBGColor(lcl_bgcolor, "#eeeeee", "#ffffff")

        rs.movenext
     wend
  end if

  rs.close
  set rs = nothing
end sub
%>