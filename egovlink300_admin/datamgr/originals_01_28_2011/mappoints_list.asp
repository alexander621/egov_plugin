<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="mappoints_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: mappoints_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the entries in the Blog
'
' MODIFICATION HISTORY
' 1.0 03/05/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Determine if the parent feature is "offline"
 if isFeatureOffline("mappoints") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_feature            = "mappoints_maint"
 lcl_showsearchcriteria = true

 if request("f") <> "" then
    lcl_feature            = request("f")
    lcl_showsearchcriteria = false

   'Build return parameters
    lcl_url_parameters = ""
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
    response.redirect sLevel & "permissiondenied.asp"
 end if

 lcl_featurename     = getFeatureName(lcl_feature)
 lcl_mappoint_typeid = getMapPointTypeByFeature(session("orgid"), "feature_maintain", lcl_feature)

 'lcl_pagetitle = "Maintain Map Points"
 lcl_pagetitle = lcl_featurename
 lcl_success   = request("success")

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for org features
 lcl_orghasfeature_feature                 = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain        = orghasfeature(lcl_feature)
 lcl_orghasfeature_customreports_mappoints = orghasfeature("customreports_mappoints")

'Check for user permissions
 lcl_userhaspermission_feature          = userhaspermission(session("userid"),lcl_feature)
 lcl_userhaspermission_feature_maintain = userhaspermission(session("userid"),lcl_feature)

'Retrieve the search options
 lcl_sc_mappoint_typeid = ""
' lcl_sc_fromcreatedate = ""
' lcl_sc_tocreatedate   = ""
' lcl_sc_title          = ""
' lcl_sc_userid         = 0
' lcl_sc_orderby        = "createdate"

 if request("sc_mappoint_typeid") <> "" then
    lcl_sc_mappoint_typeid = request("sc_mappoint_typeid")
 end if

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

  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_mappoint_typeid", lcl_sc_mappoint_typeid)
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
  lcl_mappoint = document.getElementById("mappoint"+p_id).innerHTML;

 	if (confirm("Are you sure you want to delete '" + lcl_mappoint + "' ?")) { 
  				//DELETE HAS BEEN VERIFIED

      <%
        lcl_delete_mappoint = lcl_url_parameters
        lcl_delete_mappoint = setupUrlParameters(lcl_delete_mappoint, "user_action", "DELETE")
      %>

		  		location.href='mappoints_action.asp<%=lcl_delete_mappoint%>&mappointid='+ p_id;
		}
}

function validateFields() {
  var lcl_false_count = 0;
		var daterege        = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
		var dateFromOk      = daterege.test(document.getElementById("sc_fromcreatedate").value);
		var dateToOk        = daterege.test(document.getElementById("sc_tocreatedate").value);

  if (document.getElementById("sc_tocreatedate").value!="") {
   		if (! dateToOk ) {
         document.getElementById("sc_tocreatedate").focus();
         inlineMsg(document.getElementById("toDateCalPop").id,'<strong>Invalid Value: </strong> The "To Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'toDateCalPop');
         lcl_false_count = lcl_false_count + 1;
     }else{
         clearMsg("toDateCalPop");
     }
  }

  if (document.getElementById("sc_fromcreatedate").value!="") {
   		if (! dateFromOk ) {
         document.getElementById("sc_fromcreatedate").focus();
         inlineMsg(document.getElementById("fromDateCalPop").id,'<strong>Invalid Value: </strong> The "From Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'fromDateCalPop');
         lcl_false_count = lcl_false_count + 1;
     }else{
         clearMsg("fromDateCalPop");
     }
  }

  if(lcl_false_count > 0) {
     return false;
  }else{
     document.getElementById("searchMayorsBlog").submit();
     return true;
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

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;">
            <p>
            <table border="0" cellspacing="0" cellpadding="0" width="1000px">
              <form name="mappoints" id="mappoints" action="mappoints_list.asp" method="post">
              <tr>
                  <td><font size="+1"><strong><%=lcl_pagetitle%></strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;">&nbsp;</span></td>
              </tr>
        <% if lcl_showsearchcriteria then %>
              <tr>
                  <td colspan="2">
                      <fieldset>
                        <legend>Search Options&nbsp;</legend>
                        <p>
                        <table border="0" cellspacing="0" cellpadding="0">
                          <tr>
                              <td>
                                  Category:
                                  <select name="sc_mappoint_typeid" id="sc_mappoint_typeid">
                                    <option value=""></option>
                                    <% displayMapPointTypes session("orgid"), lcl_sc_mappoint_typeid, lcl_feature %>
                                  </select>
                              </td>
                          </tr>
                        </table>
                        </p>
                        <input type="submit" name="searchButton" id="searchButton" value="Search" class="button" />
                      </fieldset>
                  </td>
              </tr>
         <%
           end if

           response.write "  </form>" & vbcrlf
           response.write "</table>" & vbcrlf

          'If a Map-Point Type has NOT been created then do NOT allow a Map-Point to be added.
           if lcl_mappoint_typeid > 0 OR lcl_showsearchcriteria then
              response.write "<p>" & vbcrlf
              response.write "   <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
              response.write "     <tr valign=""top"">" & vbcrlf
              response.write "         <td><input type=""button"" name=""addButton"" id=""addButton"" value=""Add Map-Point"" class=""button"" onclick=""location.href='mappoints_maint.asp" & lcl_url_parameters & "'"" /></td>" & vbcrlf

              if lcl_orghasfeature_customreports_mappoints then
                 response.write "         <td align=""right""><input type=""button"" name=""exportButton"" id=""exportButton"" value=""Export Map-Points"" class=""button"" onclick=""openCustomReports('MAPPOINTS_EXPORT');"" /></td>" & vbcrlf
              end if

              response.write "     </tr>" & vbcrlf
              response.write "   </table>" & vbcrlf
              response.write "   </p>" & vbcrlf
           end if

           displayMapPoints session("orgid"), lcl_feature, lcl_mappoint_typeid, lcl_sc_mappoint_typeid, lcl_url_parameters
         %>
            </p>
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
sub displayMapPoints(p_orgid, p_feature, p_mappoint_typeid, p_sc_mappoint_typeid, p_url_parameters)
 	Dim iRowCount

  sSQL = "SELECT mp.mappointid, "
  sSQL = sSQL & " mp.mappoint_typeid, "
  sSQL = sSQL & " mpt.description, "
  'sSQL = sSQL & " mp.status_id, "
  'sSQL = sSQL & " es.status_name, "
  sSQL = sSQL & " mp.createdbyid, "
  sSQL = sSQL & " mp.createdbydate, "
  sSQL = ssQL & " mp.lastmodifiedbyid, "
  sSQL = sSQL & " mp.lastmodifiedbydate, "
  sSQL = sSQL & " mp.isActive, "
  sSQL = sSQL & " mp.streetnumber, "
  sSQL = sSQL & " mp.streetprefix, "
  sSQL = sSQL & " mp.streetaddress, "
  sSQL = sSQL & " mp.streetsuffix, "
  sSQL = sSQL & " mp.streetdirection, "
  sSQL = sSQL & " mp.sortstreetname, "
  sSQL = sSQL & " mp.city, "
  sSQL = sSQL & " mp.state, "
  sSQL = sSQL & " mp.zip, "
  sSQL = sSQL & " mp.latitude, "
  sSQL = sSQL & " mp.longitude "
  'sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS createdbyname, "
  'sSQL = sSQL & " u3.firstname + ' ' + u3.lastname AS lastmodifiedbyname, "
'  sSQL = sSQL & " mp.contact_firstname, "
'  sSQL = sSQL & " mp.contact_lastname, "
'  sSQL = sSQL & " mp.contact_businessname, "
'  sSQL = sSQL & " mp.contact_email, "
'  sSQL = sSQL & " mp.contact_daytime_phone "
  sSQL = sSQL & " FROM egov_mappoints mp "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_mappoints_types mpt ON mp.mappoint_typeid = mpt.mappoint_typeid "

  if p_feature <> "" AND p_feature <> "mappoints_maint" then
     sSQL = sSQL & " AND UPPER(mpt.feature_maintain) = '" & UCASE(p_feature) & "' "
  end if

  'sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON mp.createdbyid = u2.userid AND u2.orgid = " & p_orgid
  'sSQL = sSQL &      " LEFT OUTER JOIN users u3 ON mp.lastmodifiedbyid = u3.userid AND u3.orgid = " & p_orgid
'  sSQL = sSQL &      " LEFT OUTER JOIN egov_statuses es ON mp.status_id = es.status_id "
'  sSQL = sSQL &                  " AND UPPER(es.status_type) = 'MAPPOINT_" & UCASE(p_mappointtype) & "' "
'  sSQL = sSQL &                  " AND es.orgid = 0 "
  sSQL = sSQL & " WHERE mp.orgid = " & p_orgid

  if p_feature <> "" AND p_feature <> "mappoints_maint" then
     sSQL = sSQL & " AND mp.mappoint_typeid = " & p_mappoint_typeid
  else
    'Setup the WHERE clause with the search option values.
     if trim(p_sc_mappoint_typeid) <> "" then
        sSQL = sSQL & " AND mp.mappoint_typeid = " & p_sc_mappoint_typeid
     end if
  end if

  sSQL = sSQL & " ORDER BY mpt.description "
'  if trim(p_sc_fromcreatedate) <> "" then
'     sSQL = sSQL & " AND b.createdbydate >= CAST('" & p_sc_fromcreatedate & "' as datetime) "
'  end if

'  if trim(p_sc_tocreatedate) <> "" then
'     sSQL = sSQL & " AND b.createdbydate <= CAST('" & p_sc_tocreatedate & "' as datetime) "
'  end if

'  if trim(p_sc_userid) <> "" AND p_sc_userid > 0 then
'     sSQL = sSQL & " AND b.userid = " & p_sc_userid
'  end if

'  if trim(p_sc_title) <> "" then
'     sSQL = sSQL & " AND UPPER(b.title) LIKE ('%" & UCASE(p_sc_title) & "%') "
'  end if

 'Setup the ORDER BY
'  lcl_orderby = "b.createdbydate DESC"

'  if trim(p_sc_orderby) <> "" then
'     lcl_sc_orderby = trim(UCASE(p_sc_orderby))

'     if lcl_sc_orderby = "BLOGOWNER" then
'        lcl_orderby = "u.lastname, u.firstname, b.createdbydate DESC"
'     elseif lcl_sc_orderby = "CREATEDBY" then
'        lcl_orderby = "u2.lastname, u2.firstname, b.createdbydate DESC"
'     elseif lcl_sc_orderby = "ACTIVE" then
'        lcl_orderby = "b.isActive DESC, b.createdbydate DESC"
'     end if
'  end if

'  sSQL = sSQL & " ORDER BY " & lcl_orderby

  session("CR_MAPPOINTS_EXPORT") = sSQL

 	set oMapPoints = Server.CreateObject("ADODB.Recordset")
	 oMapPoints.Open sSQL, Application("DSN"), 3, 1
	
 	if not oMapPoints.eof then
   		response.write "<div class=""shadow"">" & vbcrlf
 		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:1000px"">" & vbcrlf
   		response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"">Category</th>" & vbcrlf
     response.write "      <th align=""left"">Property Address</th>" & vbcrlf
     response.write "      <th align=""left"">Latitude</th>" & vbcrlf
     response.write "      <th align=""left"">Longitude</th>" & vbcrlf
     response.write "      <th>Active</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     lcl_bgcolor             = "#ffffff"
     lcl_original_categoryid = 0

     do while not oMapPoints.eof
        lcl_bgcolor  = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     			iRowCount    = iRowCount + 1

       'Setup the onclick
        lcl_row_onclick  = setupUrlParameters(p_url_parameters, "mappointid", oMapPoints("mappointid"))
        lcl_row_onclick  = "location.href='mappoints_maint.asp" & lcl_row_onclick & "';"
        'lcl_row_onclick  = "location.href='mappoints_maint.asp?mappointid=" & oMapPoints("mappointid") & replace(p_url_parameters,"?","&") & "';"

       'Build the Address
        lcl_displayAddress = buildStreetAddress(oMapPoints("streetnumber"), oMapPoints("streetprefix"), oMapPoints("streetaddress"), oMapPoints("streetsuffix"), oMapPoints("streetdirection"))

       'Build the "active" display value
        lcl_display_active = "&nbsp;"

        if oMapPoints("isActive") then
           lcl_display_active = "Y"
        end if

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """>" & oMapPoints("description") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """><span id=""mappoint" & oMapPoints("mappointid") & """>" & lcl_displayAddress & "</span></td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """>" & oMapPoints("latitude")  & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """>" & oMapPoints("longitude") & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" title=""click to edit"" onclick=""" & lcl_row_onclick & """ align=""center"">" & lcl_display_active      & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" align=""center""><input type=""button"" name=""delete" & iRowCount & """ id=""delete"   & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirmDelete('" & oMapPoints("mappointid") & "');"" /></td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf

        oMapPoints.movenext
     loop

   		response.write "</table>" & vbcrlf
	    response.write "</div>" & vbcrlf

  else
   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No Records Available.</p>" & vbcrlf
  end if

 	oMapPoints.close
 	set oMapPoints = nothing

end sub

'------------------------------------------------------------------------------
sub getMapPointTypeInfo(ByVal iMapPointType, ByRef lcl_description, ByRef lcl_feature, ByRef lcl_feature_maintain)

  lcl_description      = ""
  lcl_feature          = ""
  lcl_feature_maintain = ""

  if iMapPointType <> "" then
     sSQL = "SELECT description, feature, feature_maintain "
     sSQL = sSQL & " FROM egov_mappoints_types "
     sSQL = sSQL & " WHERE UPPER(mappointtype) = '" & UCASE(iMapPointType) & "' "

     set oGetMPTypeInfo = Server.CreateObject("ADODB.Recordset")
    	oGetMPTypeInfo.Open sSQL, Application("DSN"), 3, 1

     if not oGetMPTypeInfo.eof then
        lcl_description      = oGetMPTypeInfo("description")
        lcl_feature          = oGetMPTypeInfo("feature")
        lcl_feature_maintain = oGetMPTypeInfo("feature_maintain")
     end if

     oGetMPTypeInfo.close
     set oGetMPTypeInfo = nothing
  end if

end sub
%>