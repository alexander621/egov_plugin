<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_layouts_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the datamgr layouts
'
' MODIFICATION HISTORY
' 1.0 01/28/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel          = "../"  'Override of value from common.asp
 lcl_isRootAdmin = False

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Determine if the parent feature is "offline"
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user has access to maintain
 lcl_feature     = "datamgr_maintain_layouts"
 lcl_featurename = getFeatureName(lcl_feature)
 lcl_pagetitle   = lcl_featurename
 lcl_success     = request("success")

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for org features
 lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
<!--
function confirmDelete(p_id) {
  lcl_datamgr = document.getElementById("datamgr"+p_id).innerHTML;

 	if (confirm("Are you sure you want to delete '" + lcl_datamgr + "' ?")) { 
  				//DELETE HAS BEEN VERIFIED
		  		location.href='datamgr_action.asp<%=lcl_delete_datamgr%>&dmid='+ p_id;
		}
}

function doCalendar(ToFrom) {
  w = 350;
  h = 250;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0")');
}

function openCustomReports(p_report) {
  w = 900;
  h = 500;
  t = (screen.availHeight/2)-(h/2);
  l = (screen.availWidth/2)-(w/2);
  eval('window.open("../customreports/customreports.asp?cr=' + p_report + '&mpt=<%=lcl_mappoint_typeid%>", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,resizable=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
 response.write "<div id=""content"">" & vbcrlf
 response.write " 	<div id=""centercontent"">" & vbcrlf
 response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"">" & vbcrlf
 response.write "          <div style=""margin-top:20px; margin-left:20px;"">" & vbcrlf
 response.write "            <p>" & vbcrlf
 response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""1000px"">" & vbcrlf
 response.write "              <tr>" & vbcrlf
 response.write "                  <td><font size=""+1""><strong>" & lcl_pagetitle & "</strong></font></td>" & vbcrlf
 response.write "                  <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;"">&nbsp;</span></td>" & vbcrlf
 response.write "              </tr>" & vbcrlf
 response.write "            </table>" & vbcrlf
 response.write "            </p>" & vbcrlf
 response.write "            <p>" & vbcrlf
 response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
 response.write "              <tr valign=""top"">" & vbcrlf
 response.write "                  <td>" & vbcrlf
 response.write "                      <input type=""button"" name=""addButton"" id=""addButton"" value=""Add Layout"" class=""button"" onclick=""location.href='datamgr_layouts_maint.asp'"" />" & vbcrlf
 response.write "                      <input type=""button"" name=""maintainSectionsButton"" id=""maintainSectionsButton"" value=""Maintain Sections"" class=""button"" onclick=""location.href='datamgr_sections_list.asp'"" />" & vbcrlf
 response.write "                  </td>" & vbcrlf
 response.write "              </tr>" & vbcrlf
 response.write "            </table>" & vbcrlf
 response.write "            </p>" & vbcrlf

                             displayDMLayouts session("orgid")

 response.write "          </div>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"--> 
<%
 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayDMLayouts(p_orgid)
 	Dim iRowCount

  sSQL = "SELECT layoutid, "
  sSQL = sSQL & " layoutname, "
  sSQL = sSQL & " totalcolumns, "
  sSQL = sSQL & " isActive, "
  sSQL = sSQL & " useLayoutSections "
  sSQL = sSQL & " FROM egov_dm_layouts "
  sSQL = sSQL & " ORDER BY isActive DESC, layoutname "

 	set oDMLayouts = Server.CreateObject("ADODB.Recordset")
	 oDMLayouts.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMLayouts.eof then
     lcl_bgcolor = "#ffffff"

   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:800px"">" & vbcrlf
   		response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"">Layout</th>" & vbcrlf
     response.write "      <th>Total<br />Columns</th>" & vbcrlf
     response.write "      <th>""Sections""<br />Enabled</th>" & vbcrlf
     response.write "      <th>Active</th>" & vbcrlf
     'response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     do while not oDMLayouts.eof
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     			iRowCount    = iRowCount + 1

       'Setup the onclick
        lcl_row_onclick  = "location.href='datamgr_layouts_maint.asp?layoutid=" & oDMLayouts("layoutid") & "';"

       'Build the "active" display value
        lcl_display_active   = "&nbsp;"
        lcl_display_sections = "&nbsp;"

        if oDMLayouts("isActive") then
           lcl_display_active = "Y"
        end if

        if oDMLayouts("useLayoutSections") then
           lcl_display_sections = "Y"
        end if

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMLayouts("layoutname") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & oDMLayouts("totalcolumns") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & lcl_display_sections       & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & lcl_display_active         & "</td>" & vbcrlf
        'response.write "      <td class=""formlist"" align=""center""><input type=""button"" name=""delete" & iRowCount & """ id=""delete"   & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirmDelete('" & oDMLayouts("layoutid") & "');"" /></td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

        oDMLayouts.movenext
     loop

   		response.write "</table>" & vbcrlf
	    response.write "</div>" & vbcrlf

  else
   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No Records Available.</p>" & vbcrlf
  end if

 	oDMLayouts.close
 	set oDMLayouts = nothing

end sub
%>