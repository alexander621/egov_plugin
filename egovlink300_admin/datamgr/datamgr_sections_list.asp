<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_sections_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the datamgr sections to be used in datamgr layouts
'
' MODIFICATION HISTORY
' 1.0 02/01/2011 David Boyer - Initial Version
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
 lcl_feature     = "datamgr_maintain_sections"
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

'Check for search options
 lcl_sc_section_orgid = ""

 if request("sc_section_orgid") <> "" then
    lcl_sc_section_orgid = request("sc_section_orgid")
 end if
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
  eval('window.open("../customreports/customreports.asp?cr=' + p_report + '&dmt=<%=lcl_dm_typeid%>", "_customreports", "width='+w+',height='+h+',toolbar=0,statusbar=0,scrollbars=1,resizable=1,menubar=0,left=' + l + ',top=' + t + '")');
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
 response.write "<form name=""sections_list"" id=""sections_list"" method=""post"" action=""datamgr_sections_list.asp"">" & vbcrlf
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
 response.write "              <fieldset class=""fieldset"">" & vbcrlf
 response.write "                <legend>Search Options</legend>" & vbcrlf
 response.write "                <table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
 response.write "                  <tr>" & vbcrlf
 response.write "                      <td>" & vbcrlf
 response.write "                          Organization:" & vbcrlf
 response.write "                          <select name=""sc_section_orgid"" id=""sc_section_orgid"">" & vbcrlf
 response.write "                            <option value=""""></option>" & vbcrlf
                                             displayOrgSearchOptions lcl_sc_section_orgid
 response.write "                          </select>" & vbcrlf
 response.write "                      </td>" & vbcrlf
 response.write "                  </tr>" & vbcrlf
 response.write "                  <tr>" & vbcrlf
 response.write "                      <td>" & vbcrlf
 response.write "                          <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
 response.write "                      </td>" & vbcrlf
 response.write "                  </tr>" & vbcrlf
 response.write "                </table>" & vbcrlf
 response.write "              </fieldset>" & vbcrlf
 response.write "            </p>" & vbcrlf
 response.write "            <p>" & vbcrlf
 response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
 response.write "              <tr valign=""top"">" & vbcrlf
 response.write "                  <td>" & vbcrlf
 response.write "                      <input type=""button"" name=""addButton"" id=""addButton"" value=""Add Section"" class=""button"" onclick=""location.href='datamgr_sections_maint.asp'"" />" & vbcrlf
 response.write "                      <input type=""button"" name=""maintainLayoutsButton"" id=""maintainLayoutsButton"" value=""Maintain Layouts"" class=""button"" onclick=""location.href='datamgr_layouts_list.asp'"" />" & vbcrlf

 if lcl_isRootAdmin then
    response.write "                      <input type=""button"" name=""transferFieldDataButton"" id=""transferFieldDataButton"" value=""Transfer Field Data"" class=""button"" onclick=""location.href='datamgr_transfer_field_data.asp'"" />" & vbcrlf
 end if

 response.write "                  </td>" & vbcrlf
 response.write "              </tr>" & vbcrlf
 response.write "            </table>" & vbcrlf
 response.write "            </p>" & vbcrlf

                             displayDataMgrSections lcl_sc_section_orgid

 response.write "          </div>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "  </div>" & vbcrlf
 response.write "</div>" & vbcrlf
 response.write "</form>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"--> 
<%
 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayDataMgrSections(p_sc_section_orgid)
 	dim iRowCount, lcl_sc_sorgid

  lcl_sc_sorgid = 0

  if p_sc_section_orgid <> "" then
     lcl_sc_sorgid = clng(p_sc_section_orgid)
  end if

  sSQL = "SELECT dms.sectionid, "
  sSQL = sSQL & " dms.sectionname, "
  sSQL = sSQL & " dms.sectiontype, "
  sSQL = sSQL & " dms.description, "
  sSQL = sSQL & " dms.isActive, "
  sSQL = sSQL & " dms.section_orgid, "
  sSQL = sSQL & " o.orgcity "
  sSQL = sSQL & " FROM egov_dm_sections dms "
  sSQL = sSQL &      " LEFT OUTER JOIN organizations o ON dms.section_orgid = o.orgid "

  if lcl_sc_sorgid > 0 then
     sSQL = sSQL & " WHERE dms.section_orgid = " & lcl_sc_sorgid
  end if

  sSQL = sSQL & " ORDER BY o.orgcity, dms.sectionname "

 	set oDMSections = Server.CreateObject("ADODB.Recordset")
	 oDMSections.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMSections.eof then
     lcl_bgcolor = "#ffffff"

   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:800px"">" & vbcrlf
   		response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"">Section</th>" & vbcrlf
     response.write "      <th align=""left"">Description</th>" & vbcrlf
     response.write "      <th align=""left"">Org</th>" & vbcrlf
     response.write "      <th>Active</th>" & vbcrlf
     'response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     do while not oDMSections.eof
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     			iRowCount    = iRowCount + 1

       'Setup the onclick
        lcl_row_onclick  = "location.href='datamgr_sections_maint.asp?sectionid=" & oDMSections("sectionid") & "';"

       'Build the "active" display value
        lcl_display_active      = "&nbsp;"
        lcl_display_description = "&nbsp;"
        lcl_section_orgid       = 0
        lcl_section_orgcity     = "&nbsp;"

        if oDMSections("isActive") then
           lcl_display_active = "Y"
        end if

        if oDMSections("description") <> "" then
           lcl_display_description = oDMSections("description")
        end if

        if oDMSections("section_orgid") <> "" then
           lcl_section_orgid   = oDMSections("section_orgid")
           lcl_section_orgcity = oDMSections("orgcity")
        end if

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMSections("sectionname") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & lcl_display_description    & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & lcl_section_orgcity        & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & lcl_display_active & "</td>" & vbcrlf
        'response.write "      <td class=""formlist"" align=""center""><input type=""button"" name=""delete" & iRowCount & """ id=""delete"   & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirmDelete('" & oDMSections("layoutid") & "');"" /></td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

        oDMSections.movenext
     loop

   		response.write "</table>" & vbcrlf
	    response.write "</div>" & vbcrlf

  else
   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No Records Available.</p>" & vbcrlf
  end if

 	oDMSections.close
 	set oDMSections = nothing

end sub

'------------------------------------------------------------------------------
sub displayOrgSearchOptions(iSC_SectionOrgID)

  dim sSQL, sOrgID, sSC_SectionOrgID

  sSC_SectionOrgID = 0

  if iSC_SectionOrgID <> "" then
     sSC_SectionOrgID = clng(iSC_SectionOrgID)
  end if

  sSQL = "SELECT distinct o.orgcity, "
  sSQL = sSQL & " s.section_orgid "
  sSQL = sSQL & " FROM egov_dm_sections s "
  sSQL = sSQL &      " INNER JOIN organizations o ON s.section_orgid = o.orgid "
  sSQL = sSQL & " ORDER BY o.orgcity "

 	set oSCOrgOptions = Server.CreateObject("ADODB.Recordset")
	 oSCOrgOptions.Open sSQL, Application("DSN"), 3, 1

  if not oSCOrgOptions.eof then
     do while not oSCOrgOptions.eof

        if sSC_SectionOrgID = clng(oSCOrgOptions("section_orgid")) then
           lcl_selected_org = " selected=""selected"""
        else
           lcl_selected_org = ""
        end if

        response.write "  <option value=""" & oSCOrgOptions("section_orgid") & """" & lcl_selected_org & ">" & oSCOrgOptions("orgcity") & " [" & oSCOrgOptions("section_orgid") & "]</option>" & vbcrlf

        oSCOrgOptions.movenext
     loop
  end if

  oSCOrgOptions.close
  set oSCOrgOptions = nothing

end sub
%>