<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: orgs_to_feature_summary.asp
' AUTHOR: Steve Loar
' CREATED: 10/15/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of all orgs that have a specific feature
'
' MODIFICATION HISTORY
' 1.0  07/01/2011  David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

 if not userIsRootAdmin(session("UserID")) then
   	response.redirect "../default.asp"
 end if

'Check for search options
 lcl_sc_featureid = ""

 if request("sc_featureid") <> "" then
    lcl_sc_featureid = clng(request("sc_featureid"))
 end if
%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<title>E-GovLink {Orgs-to-Feature Summary}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="admin.css" />

<style type="text/css">
  #search_fieldset {
     border:                1pt solid #808080;
     -moz-border-radius:    5px;
     -webkit-border-radius: 5px;
  }
</style>

</head>
<body>
 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "		<div id=""centercontent"">" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "				<font size=""+1""><strong>Orgs-to-Feature Summary</strong></font><br />" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "<div class=""filterselection"">" & vbcrlf
  response.write "  <fieldset id=""search_fieldset"">" & vbcrlf
  response.write "				<legend class=""filterselection"">Search Options</legend>" & vbcrlf
  response.write "				<p>" & vbcrlf
  response.write "				  <form name=""searchoptions"" id=""searchoptions"" method=""post"" action=""orgs_to_feature_summary.asp"">" & vbcrlf
  response.write "						<table cellpadding=""2"" cellspacing=""0"" border=""0"" id=""pagelogpicks"">" & vbcrlf
  response.write "						  <tr>" & vbcrlf
  response.write "								    <td>Feature:</td>" & vbcrlf
  response.write "            <td>" & vbcrlf
  response.write "                <select name=""sc_featureid"" id=""sc_featureid"">" & vbcrlf
                                    displayFeatureOptions lcl_sc_featureid
  response.write "                </select>" & vbcrlf
  response.write "            </td>" & vbcrlf
  response.write "								</tr>" & vbcrlf
  response.write "								<tr>" & vbcrlf
  response.write "								 			<td colspan=""2"">" & vbcrlf
  response.write "                <input type=""submit"" class=""button"" value=""Search"" onclick=""RefreshResults();"" />" & vbcrlf
  response.write "            </td>" & vbcrlf
  response.write "								</tr>" & vbcrlf
  response.write "						</table>" & vbcrlf
  response.write "						</form>" & vbcrlf
  response.write "				</p>" & vbcrlf
  response.write "		</fieldset>" & vbcrlf
  response.write "</div>" & vbcrlf
                  displayOrganizationsList lcl_sc_featureid
  response.write "</div>" & vbcrlf
  response.write "		</div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
	<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayFeatureOptions(iSCFeatureID)

  sSCFeatureID = ""

  if iSCFeatureID <> "" then
     sSCFeatureID = clng(iSCFeatureID)
  end if

  response.write "  <option value=""0"">&nbsp;</option>" & vbcrlf

  sSQL = "SELECT f.parentfeatureid, "
  sSQL = sSQL & "(select f2.featurename "
  sSQL = sSQL &  "from egov_organization_features f2 "
  sSQL = sSQL &  "where f2.featureid = f.parentfeatureid) parentfeaturename, "
  sSQL = sSQL & "f.featureid, "
  sSQL = sSQL & "f.featurename "
  sSQL = sSQL & " FROM egov_organization_features f "
  sSQL = sSQL & " ORDER BY (select upper(f2.featurename) "
  sSQL = sSQL &           "from egov_organization_features f2 "
  sSQL = sSQL &           "where f2.featureid = f.parentfeatureid), "
  sSQL = sSQL & " upper(f.featurename) "

 	set oFeatureOptions = Server.CreateObject("ADODB.Recordset")
 	oFeatureOptions.Open sSQL, Application("DSN"), 0, 1

  if not oFeatureOptions.eof then
     do while not oFeatureOptions.eof
        lcl_selected_featureid = ""
        lcl_parent_featurename = ""

        if sSCFeatureID = oFeatureOptions("featureid") then
           lcl_selected_featureid = " selected=""selected"""
        end if

        if oFeatureOptions("parentfeaturename") <> "" then
           lcl_parent_featurename = oFeatureOptions("parentfeaturename") & ": "
        end if

        response.write "  <option value=""" & oFeatureOptions("featureid") & """" & lcl_selected_featureid & ">" & lcl_parent_featurename & oFeatureOptions("featurename") & "</option>" & vbcrlf

        oFeatureOptions.movenext
     loop
  end if

  oFeatureOptions.close
  set oFeatureOptions = nothing

end sub

'------------------------------------------------------------------------------
sub displayOrganizationsList(iSCFeatureID)

  sSCFeatureID   = 0
  lcl_bgcolor    = "#eeeeee"
  lcl_line_count = 0

  if iSCFeatureID <> "" then
     sSCFeatureID = clng(iSCFeatureID)
  end if

  sSQL = "SELECT otf.orgid, "
  sSQL = sSQL & " o.orgname "
  sSQL = sSQL & " FROM egov_organizations_to_features AS otf "
  sSQL = sSQL &      " INNER JOIN Organizations AS o ON otf.orgid = o.OrgID "
  sSQL = sSQL & " WHERE otf.featureid = " & sSCFeatureID
  sSQL = sSQL & " and o.isdeactivated = 0 "
  sSQL = sSQL & " ORDER BY o.OrgName "

 	set oOrgToFeatureList = Server.CreateObject("ADODB.Recordset")
 	oOrgToFeatureList.Open sSQL, Application("DSN"), 0, 1

  if not oOrgToFeatureList.eof then

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"">" & vbcrlf
     response.write "  <tr align=""left"">" & vbcrlf
     response.write "      <th>Organizations</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     do while not oOrgToFeatureList.eof
        lcl_line_count = lcl_line_count + 1
        lcl_bgcolor    = changeBGColor(lcl_bgcolor, "#eeeeee", "#ffffff")

        response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
        response.write "      <td>" & oOrgToFeatureList("orgname") & " [" & oOrgToFeatureList("orgid") & "]</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oOrgToFeatureList.movenext
     loop

     response.write "</table>" & vbcrlf
     response.write "<div align=""right""><strong>Total Organizations:</strong> [" & lcl_line_count & "]</div>" & vbcrlf

  end if

  oOrgToFeatureList.close
  set oOrgToFeatureList = nothing

end sub
%>