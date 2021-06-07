<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappoints_types_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This screen allows an e-gov root admin to maintain all map point types.
'
' MODIFICATION HISTORY
' 1.0 03/05/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("mappoints") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"mappoints_types_maint") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 lcl_pagetitle = "Map-Point Categories"
 lcl_success   = request("success")

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if


'Check for org features
 lcl_orghasfeature_mappoints             = orghasfeature("mappoints")
 lcl_orghasfeature_mappoints_types_maint = orghasfeature("mappoint_types_maint")

'Check for user permissions
 lcl_userhaspermission_mappoints             = userhaspermission(session("userid"),"mappoints")
 lcl_userhaspermission_mappoints_types_maint = userhaspermission(session("userid"),"mappoints_types_maint")

'Retrieve the search options
' lcl_sc_fromcreatedate = ""
' lcl_sc_tocreatedate   = ""
' lcl_sc_title          = ""
' lcl_sc_userid         = 0
' lcl_sc_orderby        = "createdate"

' if request("sc_fromcreatedate") <> "" then
'    lcl_sc_fromcreatedate = request("sc_fromcreatedate")
' end if

' if request("sc_tocreatedate") <> "" then
'    lcl_sc_tocreatedate = request("sc_tocreatedate")
' end if

' if request("sc_title") <> "" then
'    lcl_sc_title = request("sc_title")
' end if

' if request("sc_userid") <> "" then
'    lcl_sc_userid = request("sc_userid")
' end if

' if request("sc_orderby") <> "" then
'    lcl_sc_orderby = request("sc_orderby")
' end if
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
function confirm_delete(iMapPointTypeID) {
  lcl_description = document.getElementById("description"+iMapPointTypeID).innerHTML;

 	if (confirm("Are you sure you want to delete: '" + lcl_description + "' ?")) { 
  				//DELETE HAS BEEN VERIFIED
		  		location.href='mappoints_types_action.asp?user_action=DELETE&mappoint_typeid='+ iMapPointTypeID;
		}
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

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;">
            <p>
            <table border="0" cellspacing="0" cellpadding="0" width="1000px">
              <tr>
                  <td><font size="+1"><strong><%=lcl_pagetitle%></strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
              </tr>
            </table>
          <%
            if lcl_userhaspermission_mappoints_types_maint then
               response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add Category"" class=""button"" onclick=""location.href='mappoints_types_maint.asp';"" />" & vbcrlf
            end if
          %>
            </p>
            <% displayMapPointTypes session("orgid") %>
          </div>
      </td>
  </tr>
</table>

  </div>
</div>
	
<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displayMapPointTypes(p_orgid)
 	Dim iRowCount

 	iRowCount = 0

  sSQL = "SELECT mpt.mappoint_typeid, "
  sSQL = sSQL & " mpt.isInactive, "
  sSQL = sSQL & " mpt.description, "
  sSQL = sSQL & " mpt.createdbyid, mpt.createdbydate, "
  sSQL = sSQL & " (select u.firstname + ' ' + u.lastname "
  sSQL = sSQL &   "from users u "
  sSQL = sSQL &   "where u.userid = mpt.createdbyid) as createdbyname, "
  sSQL = sSQL & " mpt.lastmodifiedbyid, "
  sSQL = sSQL & " mpt.lastmodifiedbydate, "
  sSQL = sSQL & " (select u2.firstname + ' ' + u2.lastname "
  sSQL = sSQL &   "from users u2 "
  sSQL = sSQL &   "where u2.userid = mpt.lastmodifiedbyid) as lastmodifiedbyname "
  sSQL = sSQL & " FROM egov_mappoints_types mpt "
  sSQL = sSQL & " WHERE orgid = " & p_orgid
  sSQL = sSQL & " ORDER BY mpt.description "

 'Setup the WHERE clause with the search option values.
 ' if trim(p_sc_fromcreatedate) <> "" then
 '    sSQL = sSQL & " AND b.createdbydate >= CAST('" & p_sc_fromcreatedate & "' as datetime) "
 ' end if

 ' if trim(p_sc_tocreatedate) <> "" then
 '    sSQL = sSQL & " AND b.createdbydate <= CAST('" & p_sc_tocreatedate & "' as datetime) "
 ' end if

 ' if trim(p_sc_userid) <> "" AND p_sc_userid > 0 then
 '    sSQL = sSQL & " AND b.userid = " & p_sc_userid
 ' end if

 ' if trim(p_sc_title) <> "" then
 '    sSQL = sSQL & " AND UPPER(b.title) LIKE ('%" & UCASE(p_sc_title) & "%') "
 ' end if

 	set oMapPointTypes = Server.CreateObject("ADODB.Recordset")
	 oMapPointTypes.Open sSQL, Application("DSN"), 3, 1
	
 	if not oMapPointTypes.eof then
   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:800px"">" & vbcrlf
   		response.write "  <tr align=""left"">" & vbcrlf
     response.write "      <th>Description</th>" & vbcrlf
     response.write "      <th align=""center"">Active</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"" align=""center"">Created By</th>" & vbcrlf
     response.write "      <th nowrap=""nowrap"" align=""center"">Last Modified By</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     lcl_bgcolor             = "#ffffff"
     lcl_original_categoryid = 0

     do while not oMapPointTypes.eof
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
    	 		iRowCount    = iRowCount + 1

       'Setup the date info
        lcl_display_createdbydate      = ""
        lcl_display_lastmodifiedbydate = ""

        if trim(oMapPointTypes("createdbydate")) <> "" then
           lcl_display_createdbydate = "[" & oMapPointTypes("createdbydate") & "]"
        end if

        if trim(oMapPointTypes("lastmodifiedbydate")) <> "" then
           lcl_display_modifiedbybydate = "[" & oMapPointTypes("lastmodifiedbydate") & "]"
        end if

       'Check to see if this Map-Point Type is active or not
        if oMapPointTypes("isInactive") then
           lcl_isInactive = ""
        else
           lcl_isInactive = "Y"
        end if

       'Setup the onclick
        lcl_row_onclick = "location.href='mappoints_types_maint.asp?mappoint_typeid=" & oMapPointTypes("mappoint_typeid") & "';"

       'Check for associated Map-Points to determine if this MapPointType can/cannot be deleted.
        lcl_canDelete = checkForMapPointsByMapPointTypeID(oMapPointTypes("mappoint_typeid"))

        if lcl_canDelete then
           lcl_delete_row = "<input type=""button"" name=""delete" & iRowCount & """ id=""delete" & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & oMapPointTypes("mappoint_typeid") & "');"" />"
        else
           lcl_delete_row = "&nbsp;"
        end if

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """><span id=""description" & oMapPointTypes("mappoint_typeid") & """>" & oMapPointTypes("description") & "</span></td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" align=""center"" onclick=""" & lcl_row_onclick & """>" & lcl_isInactive & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" align=""center"">" & lcl_delete_row & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """ width=""150"" nowrap=""nowrap"" align=""center"">" & vbcrlf
        response.write            trim(oMapPointTypes("createdbyname")) & "<br />" & vbcrlf
        response.write "          <span style=""color:#800000;"">" & lcl_display_createdbydate & "</span>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """ width=""150"" nowrap=""nowrap"" align=""center"">" & vbcrlf
        response.write            trim(oMapPointTypes("lastmodifiedbyname")) & "<br />" & vbcrlf
        response.write "          <span style=""color:#800000;"">" & lcl_display_lastmodifiedbydate & "</span>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

        oMapPointTypes.movenext
     loop

   		response.write "</table>" & vbcrlf
  	  response.write "</div>" & vbcrlf

  else
   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No Map-Point Categories have been created.</p>" & vbcrlf
	 end if

 	oMapPointTypes.close
 	set oMapPointTypes = nothing 

end sub
%>