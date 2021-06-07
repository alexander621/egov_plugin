<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the entries in the DM Type
'
' MODIFICATION HISTORY
' 1.0 03/05/10 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Determine if the parent feature is "offline"
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_isRootAdmin        = false
 lcl_feature            = "datamgr_maint"
 lcl_showsearchcriteria = true

 if request("f") <> "" AND request("f") <> "datamgr_maint" then
    lcl_feature            = request("f")
    lcl_showsearchcriteria = false

   'Build return parameters
    lcl_url_parameters = ""
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
    response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = true
 end if

 lcl_featurename = getFeatureName(lcl_feature)
 lcl_dm_typeid   = getDMTypeByFeature(session("orgid"), "feature_maintain", lcl_feature)

 'lcl_pagetitle = "Maintain Map Points"
 lcl_pagetitle = lcl_featurename
 lcl_success   = request("success")

'Check for a screen message
 lcl_onload = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Set up approval status message
 lcl_onload = lcl_onload & "checkWaitingForApproval();"

'Check for org features
 lcl_orghasfeature_feature               = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain      = orghasfeature(lcl_feature)
 lcl_orghasfeature_customreports_datamgr = orghasfeature("customreports_datamgr")

'Check for user permissions
 lcl_userhaspermission_feature          = userhaspermission(session("userid"),lcl_feature)
 lcl_userhaspermission_feature_maintain = userhaspermission(session("userid"),lcl_feature)

'Retrieve the search options
 lcl_sc_dm_typeid      = ""
 lcl_sc_approvedDenied = ""
 lcl_sc_dm_importid    = ""
' lcl_sc_fromcreatedate = ""
' lcl_sc_tocreatedate   = ""
' lcl_sc_title          = ""
' lcl_sc_userid         = 0
' lcl_sc_orderby        = "createdate"
 if request("sc_dm_typeid") <> "" then
    lcl_sc_dm_typeid = request("sc_dm_typeid")
    lcl_sc_dm_typeid = clng(lcl_sc_dm_typeid)
 end if

 if request("sc_approvedDenied") <> "" then
    lcl_sc_approvedDenied = request("sc_approvedDenied")
 end if

 if request("sc_dm_importid") <> "" then
    lcl_sc_dm_importid = request("sc_dm_importid")
    lcl_sc_dm_importid = clng(lcl_sc_dm_importid)
 end if

 if request("sc_searchfield_0") <> "" then
    lcl_sc_searchfield = request("sc_searchfield_0")
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

  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_dm_typeid",      lcl_sc_dm_typeid)
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_approvedDenied", lcl_sc_approvedDenied)
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_dm_importid",    lcl_sc_dm_importid)

  session("RedirectPage") = ""
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
  <link rel="stylesheet" type="text/css" href="layout_styles.css" />

<style type="text/css">
  .searchLabel {
     white-space: nowrap;
  }
</style>

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

 	<script type="text/javascript" src="../scripts/column_sorting.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>
<% '  <script type="text/javascript" src="https://github.com/jquery/jquery-ui.git"></script> %>

<script language="javascript">
<!--
var sorter = new TINY.table.sorter("sorter");

function listOrderInit() {
  sorter.head      = "head";
  sorter.asc       = "asc";
  sorter.desc      = "desc";
  sorter.even      = "evenrow";
  sorter.odd       = "oddrow";
  sorter.evensel   = "evenselected";
  sorter.oddsel    = "oddselected";
  //sorter.paginate  = true;
  //sorter.currentid = "currentpage";
  //sorter.limitid   = "pagelimit";
  sorter.init("mappoints",0);
}

function approveDenyDMID(iRowID, iAction) {
  var lcl_dmid       = $('#dmid' + iRowID).val();
  var lcl_isApproved = false;

  if(iAction != '') {
     if(iAction == 'A') {
        lcl_isApproved = true;
     }

     //Approve/Deny the DMID
     $.post('approveDenyDMID.asp', {
        userid:     '<%=session("userid")%>',
        orgid:      '<%=session("orgid")%>',
        dmid:       lcl_dmid,
        isApproved: lcl_isApproved,
        isAjax:     'Y'
     }, function(result) {
        var lcl_display_status = '';
        var lcl_display_info   = '';
        var lcl_status_value   = '';
        var lcl_button_nameid  = '';
        var lcl_button_value   = '';
        var lcl_button_action  = '';
        var lcl_button         = '';

        if(result == 'approved') {
           lcl_status_value  = 'APPROVED';
           lcl_button_nameid = 'denyButton' + iRowID;
           lcl_button_value  = 'Deny';
           lcl_button_action = 'D';
        } else {
           lcl_status_value  = 'DENIED';
           lcl_button_nameid = 'approveButton' + iRowID;
           lcl_button_value  = 'Approve';
           lcl_button_action = 'A';
        }

        //Build the approve/deny info
        lcl_display_status  = '<span class="redText">' + lcl_status_value + '</span><br />';

        //Build the approve/deny button
        lcl_button  = "<input ";
        lcl_button +=   "type='button' ";
        lcl_button +=   "class='button' ";
        lcl_button +=   "name='"  + lcl_button_nameid + "' ";
        lcl_button +=   "id='"    + lcl_button_nameid + "' ";
        lcl_button +=   "value='" + lcl_button_value  + "' ";
        lcl_button +=   "onclick='approveDenyDMID(\"" + iRowID + "\",\"" + lcl_button_action + "\");' ";
        lcl_button += "/>";

        $('#approvedDeniedStatus'  + iRowID).html(lcl_display_status);
        $('#approvedDeniedInfo'    + iRowID).html(lcl_display_info);
        $('#approvedDeniedButtons' + iRowID).html(lcl_button);
     });

  }
}

function checkWaitingForApproval() {
  var lcl_total_waiting = '';
  var lcl_approval_msg  = '';

  lcl_total_waiting = $('#total_waiting').val();

  if(lcl_total_waiting > 0) {
     lcl_approval_msg = '<p>*** ' + lcl_total_waiting + ' waiting for approval ***</p>';
  }

  $('#waitingApprovalText').html(lcl_approval_msg);
}

function confirmDelete(p_id) {
  //var lcl_dmdata = document.getElementById("dmdata"+p_id).innerHTML;

 	if (confirm("Are you sure you want to delete this record?")) { 
  				//DELETE HAS BEEN VERIFIED

      <%
        lcl_delete_dmdata = lcl_url_parameters
        lcl_delete_dmdata = setupUrlParameters(lcl_delete_dmdata, "user_action", "DELETE")
      %>

		  		location.href='datamgr_action.asp<%=lcl_delete_dmdata%>&dmid='+ p_id;
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
     document.getElementById("datamgr").submit();
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
  response.write "<form name=""datamgr"" id=""datamgr"" action=""datamgr_list.asp"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
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

'  if lcl_showsearchcriteria then
     lcl_dm_type_label = lcl_featurename
     lcl_dm_type_label = replace(lcl_dm_type_label,"Maintain ","")

     response.write "              <tr valign=""top"">" & vbcrlf
     response.write "                  <td colspan=""2"">" & vbcrlf
     response.write "                      <fieldset class=""fieldset"">" & vbcrlf
     response.write "                        <legend>Search Options&nbsp;</legend>" & vbcrlf
     response.write "                        <p>" & vbcrlf
     response.write "                        <table border=""0"" cellspacing=""1"" cellpadding=""2"">" & vbcrlf
     response.write "                          <tr valign=""top"">" & vbcrlf
     response.write "                              <td class=""searchLabel"">" & vbcrlf
     response.write "                                  Search:" & vbcrlf
     response.write "                              </td>" & vbcrlf
     response.write "                              <td>" & vbcrlf
     response.write "                                  <input type=""text"" name=""sc_searchfield_0"" id=""sc_searchfield_0"" size=""30"" maxlength=""30"" value=""" & lcl_sc_searchfield & """ />" & vbcrlf
     response.write "                              </td>" & vbcrlf
     response.write "                              <td class=""searchLabel"">" & vbcrlf
     response.write "                                  Approval Status:" & vbcrlf
     response.write "                                  <select name=""sc_approvedDenied"" id=""sc_approvedDenied"">" & vbcrlf
                                                         displayApprovedDeniedOptions lcl_sc_approvedDenied
     response.write "                                  </select>" & vbcrlf
     response.write "                                  <div id=""waitingApprovalText"" class=""approvalStatus""></div>" & vbcrlf
     response.write "                              </td>" & vbcrlf

     if lcl_isRootAdmin then
        response.write "                              <td class=""searchLabel"" width=""100%"">" & vbcrlf
        response.write "                                  DM ImportID:" & vbcrlf
        response.write "                                  <select name=""sc_dm_importid"" id=""sc_dm_importid"">" & vbcrlf
        response.write "                                    <option value=""""></option>" & vbcrlf
                                                            displayDMImportIDs session("orgid"), lcl_feature, lcl_sc_dm_importid
        response.write "                                  </select>" & vbcrlf
        response.write "                              </td>" & vbcrlf
     end if

     response.write "                          </tr>" & vbcrlf

     if lcl_showsearchcriteria then
        response.write "                          <tr>" & vbcrlf
        response.write "                              <td class=""searchLabel"">" & vbrlf
        response.write "                                  DM Type:" & vbcrlf
        response.write "                              </td>" & vbcrlf
        response.write "                              <td colspan=""3"">" & vbcrlf
        response.write "                                  <select name=""sc_dm_typeid"" id=""sc_dm_typeid"">" & vbcrlf
        response.write "                                    <option value=""""></option>" & vbcrlf
                                                            displayDMTypes session("orgid"), lcl_sc_dm_typeid, lcl_feature
        response.write "                                  </select>" & vbcrlf
        response.write "                              </td>" & vbcrlf
        response.write "                          </tr>" & vbcrlf
     end if

     response.write "                        </table>" & vbcrlf
     response.write "                        </p>" & vbcrlf

     if not lcl_showsearchcriteria then
        response.write "                        <input type=""hidden"" name=""sc_dm_typeid"" id=""sc_dm_typeid"" value=""" & lcl_sc_dm_typeid & """ />" & vbcrlf
     end if

     response.write "                        <input type=""submit"" name=""searchButton"" id=""searchButton"" value=""Search"" class=""button"" />" & vbcrlf
     response.write "                      </fieldset>" & vbcrlf
     response.write "                  </td>" & vbcrlf
     response.write "              </tr>" & vbcrlf
'  end if

  response.write "            </table>" & vbcrlf

 'If a DM Type has NOT been created then do NOT allow a DM Data record to be added.
  'if lcl_dm_typeid > 0 OR lcl_showsearchcriteria then
  '   lcl_addButtonLabel = lcl_featurename
  '   lcl_addButtonLabel = replace(lcl_addButtonLabel,"Maintain ","")

  '   response.write "            <p>" & vbcrlf
  '   response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  '   response.write "              <tr valign=""top"">" & vbcrlf
  '   response.write "                  <td><input type=""button"" name=""addButton"" id=""addButton"" value=""Add " & lcl_addButtonLabel & """ class=""button"" onclick=""location.href='datamgr_maint.asp" & lcl_url_parameters & "'"" /></td>" & vbcrlf
  '   response.write "                  <td align=""right"">" & vbcrlf
  '   response.write "                      <input type=""button"" name=""importFromMapPointsButton"" id=""importFromMapPointsButton"" value=""Import From MapPoints"" class=""button"" onclick=""location.href='datamgr_import_from_mappoints.asp';"" />" & vbcrlf

  '   if lcl_orghasfeature_customreports_datamgr then
  '      response.write "                  <input type=""button"" name=""exportButton"" id=""exportButton"" value=""Export Map-Points"" class=""button"" onclick=""openCustomReports('DATAMGR_EXPORT');"" />" & vbcrlf
  '   end if

  '   response.write "                  </td>" & vbcrlf
  '   response.write "              </tr>" & vbcrlf
  '   response.write "            </table>" & vbcrlf
  '   response.write "            </p>" & vbcrlf
  'end if
                              displayDMData lcl_isRootAdmin, _
                                            session("orgid"), _
                                            lcl_feature, _
                                            lcl_dm_typeid, _
                                            lcl_sc_dm_typeid, _
                                            lcl_sc_approvedDenied, _
                                            lcl_sc_dm_importid, _
                                            lcl_url_parameters, _
                                            lcl_showsearchcriteria, _
                                            lcl_orghasfeature_customreports_datamgr, _
                                            lcl_sc_searchfield

  response.write "            </p>" & vbcrlf
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
sub displayDMData(p_isRootAdmin, p_orgid, p_feature, p_dm_typeid, p_sc_dm_typeid, _
                  p_sc_approvedDenied, p_sc_dm_importid, p_url_parameters, _
                  p_showsearchcriteria, p_orghasfeature_customreports_datamgr, _
                  iSCSearchField)

 	dim iRowCount, lcl_previous_dm_typeid, lcl_waiting_count

  lcl_previous_dm_typeid = 0
  lcl_waiting_count      = 0
  lcl_scripts            = ""

  sSQL = "SELECT dm.dmid, "
  sSQL = sSQL & " dm.dm_typeid, "
  sSQL = sSQL & " dmt.description, "
  sSQL = sSQL & " dm.categoryid, "
  sSQL = sSQL & " dmc.categoryname, "
  sSQL = sSQL & " dmc.mappointcolor, "
  sSQL = sSQL & " dm.createdbyid, "
  sSQL = sSQL & " dm.createdbydate, "
  sSQL = ssQL & " dm.lastmodifiedbyid, "
  sSQL = sSQL & " dm.lastmodifiedbydate, "
  sSQL = sSQL & " dm.isActive, "
  sSQL = sSQL & " dm.streetnumber, "
  sSQL = sSQL & " dm.streetprefix, "
  sSQL = sSQL & " dm.streetaddress, "
  sSQL = sSQL & " dm.streetsuffix, "
  sSQL = sSQL & " dm.streetdirection, "
  sSQL = sSQL & " dm.sortstreetname, "
  sSQL = sSQL & " dm.city, "
  sSQL = sSQL & " dm.state, "
  sSQL = sSQL & " dm.zip, "
  sSQL = sSQL & " dm.latitude, "
  sSQL = sSQL & " dm.longitude, "
  sSQL = sSQL & " dm.isApproved, "
  sSQL = sSQL & " dm.approvedeniedbyid, "
  sSQL = sSQL & " dm.dm_importid, "
  sSQL = sSQL & " dm.approvedeniedbydate, "
  sSQL = sSQL & " dmt.accountInfoSectionID, "
  sSQL = sSQL & " u.firstname + ' ' + u.lastname AS approvedeniedbyname "
  sSQL = sSQL & " FROM egov_dm_data dm "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dm_types dmt ON dm.dm_typeid = dmt.dm_typeid "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_dm_categories dmc ON dm.categoryid = dmc.categoryid "
  sSQL = sSQL &      " LEFT OUTER JOIN users u ON dm.approvedeniedbyid = u.userid AND u.orgid = " & p_orgid

  if p_feature <> "" AND p_feature <> "datamgr_maint" then
     sSQL = sSQL & " AND UPPER(dmt.feature_maintain) = '" & UCASE(p_feature) & "' "
  end if

  sSQL = sSQL & " WHERE dm.orgid = " & p_orgid

  if p_feature <> "" AND p_feature <> "datamgr_maint" then
     sSQL = sSQL & " AND dm.dm_typeid = " & p_dm_typeid
  else
    'Setup the WHERE clause with the search option values.
     if trim(p_sc_dm_typeid) <> "" then
        sSQL = sSQL & " AND dm.dm_typeid = " & p_sc_dm_typeid
     end if
  end if

  if trim(p_sc_dm_importid) <> "" then
     sSQL = sSQL & " AND dm.dm_importid = " & p_sc_dm_importid
  end if

 'Determine if we are showing approved and/or denied DM Data
 '0 = View All
 '1 = View all WAITING for approval
 '2 = View all APPROVED
 '3 = View all DENIED
  lcl_sql_approvedDenied = ""

  if p_sc_approvedDenied <> "" then
     if p_sc_approvedDenied > 0 then
        if p_sc_approvedDenied = 1 then
           lcl_sql_approvedDenied = "= 0 AND (dm.approvedeniedbydate IS NULL OR dm.approvedeniedbydate = '') "
        elseif p_sc_approvedDenied = 2 then
           lcl_sql_approvedDenied = "= 1 "
        elseif p_sc_approvedDenied = 3 then
           lcl_sql_approvedDenied = "= 0 AND dm.approvedeniedbydate <> '' "
        end if

        sSQL = sSQL & " AND dm.isApproved " & lcl_sql_approvedDenied

     end if
  end if

  if iSCSearchField <> "" then
     lcl_fieldvalue = UCASE(iSCSearchField)
     lcl_fieldvalue = dbsafe(lcl_fieldvalue)
     lcl_fieldvalue = "'%" & lcl_fieldvalue & "%'"

     sSQL = sSQL &      " AND dm.dmid in ("
     sSQL = sSQL &                      " select distinct dmv.dmid "
     sSQL = sSQL &                      " from egov_dm_values dmv"
     sSQL = sSQL &                      " where UPPER(dmv.fieldvalue) LIKE (" & lcl_fieldvalue & ") "
     sSQL = sSQL &                      " and dmv.dm_typeid = dm.dm_typeid "

     if trim(p_sc_dm_importid) <> "" then
        sSQL = sSQL &                   " and dmv.dm_importid = " & p_sc_dm_importid
     end if

     sSQL = sSQL &                      ") "
  end if

  sSQL = sSQL & " ORDER BY dmt.description "

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

  session("CR_DATAMGR_EXPORT") = sSQL

 	set oDMData = Server.CreateObject("ADODB.Recordset")
	 oDMData.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMData.eof then
     lcl_scripts = lcl_scripts & "listOrderInit();"

'     response.write "<p>" & vbcrlf
'   		response.write "<div class=""shadow"">" & vbcrlf
'     response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" style=""width:1000px"">" & vbcrlf
'   		response.write "  <tr>" & vbcrlf

'     if p_isRootAdmin then
'        response.write "      <th align=""left"">DM Type[" & p_dm_typeid & "]</th>" & vbcrlf
'     end if

'     response.write "      <th align=""left"">Category</th>" & vbcrlf
     'response.write "      <th align=""left"">Property Address</th>" & vbcrlf
     'response.write "      <th align=""left"">Latitude</th>" & vbcrlf
     'response.write "      <th align=""left"">Longitude</th>" & vbcrlf

    'Pull all of the columns that are in the "account info" section
'     lcl_display_type = "FIELDNAME"
'     lcl_row_onclick  = ""

'     buildResultsList lcl_display_type, oDMData("dmid"), oDMData("dm_typeid"), oDMData("accountInfoSectionID"), lcl_row_onclick

'     response.write "      <th>Active</th>" & vbcrlf
'     response.write "      <th>&nbsp;</th>" & vbcrlf
'     response.write "      <th>Approved/Denied By</th>" & vbcrlf
'     response.write "      <th>&nbsp;</th>" & vbcrlf
'     response.write "  </tr>" & vbcrlf

'     lcl_bgcolor             = "#ffffff"
'     lcl_original_categoryid = 0

     do while not oDMData.eof
        lcl_bgcolor     = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_row_onclick = ""
     			iRowCount       = iRowCount + 1

       'Setup the onclick
        lcl_row_onclick  = setupUrlParameters(p_url_parameters, "dmid", oDMData("dmid"))
        lcl_row_onclick  = "location.href='datamgr_maint.asp" & lcl_row_onclick & "';"
        'lcl_row_onclick  = "location.href='datamgr_maint.asp?dmid=" & oDMData("dmid") & replace(p_url_parameters,"?","&") & "';"

       'Build the Address
        'lcl_displayAddress = buildStreetAddress(oDMData("streetnumber"), _
        '                                        oDMData("streetprefix"), _
        '                                        oDMData("streetaddress"), _
        '                                        oDMData("streetsuffix"), _
        '                                        oDMData("streetdirection"))

       'Build the "active" display value
        lcl_display_active = "&nbsp;"

        if oDMData("isActive") then
           lcl_display_active = "Y"
        end if

       'Set up Approve/Deny Buttons for display
        lcl_show_approvedButton        = 1
        lcl_show_deniedButton          = 1
        lcl_display_approvedDeniedInfo = ""
        lcl_approved_denied_status     = ""
        lcl_approvedenied_info         = formatAdminActionsInfo(oDMData("approvedeniedbyname"), oDMData("approvedeniedbydate"))

        if lcl_approvedenied_info <> "" then
           if oDMData("isApproved") then
              lcl_show_approvedButton    = 0
              lcl_approved_denied_status = "APPROVED"
           else
              lcl_show_deniedButton      = 0
              lcl_approved_denied_status = "DENIED"
           end if

           'lcl_display_approvedDeniedInfo = "<span class=""redText"">" & lcl_approved_denied_status & "</span><br />"
           'lcl_display_approvedDeniedInfo = lcl_display_approvedDeniedInfo & lcl_approvedenied_info
        else
           lcl_approved_denied_status = "WAITING FOR<br />APPROVAL"
           lcl_waiting_count          = lcl_waiting_count + 1
        end if

       'If the DM_TypeID is NOT equal to the previous DM_TypeID in the loop then close the table and open a new one.
       'The reason for this is because each DM Type has it's own "account info" fields and the columns would never line up in the results list.
       'This is mainly just for the ROOT ADMIN as clients will never see the "DM Type" column and/or search option.
        if oDMData("dm_typeid") <> lcl_previous_dm_typeid then
           if iRowCount > 1 then
            		response.write "</table>" & vbcrlf
         	    response.write "</div>" & vbcrlf
              response.write "</p>" & vbcrlf
           end if

           'if oDMData("dm_typeid") > 0 OR iShowSearchCriteria then
           'if oDMData("dm_typeid") > 0 then
              displayButtonRow p_isRootAdmin, _
                               p_feature, _
                               oDMData("dm_typeid"), _
                               oDMData("description"), _
                               lcl_url_parameters, _
                               p_orghasfeature_customreports_datamgr, _
                               lcl_url_parameters

              'lcl_addButtonLabel = oDMData("description")
              'lcl_addButtonLabel = replace(lcl_addButtonLabel,"Maintain ","")
              'lcl_addButtonURL   = "datamgr_maint.asp" & lcl_url_parameters

              'response.write "            <p>" & vbcrlf
              'response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
              'response.write "              <tr valign=""top"">" & vbcrlf
              'response.write "                  <td><input type=""button"" name=""addButton"" id=""addButton"" value=""Add " & lcl_addButtonLabel & """ class=""button"" onclick=""location.href='" & lcl_addButtonURL & "'"" /></td>" & vbcrlf
              'response.write "                  <td align=""right"">" & vbcrlf

              'if p_isRootAdmin then
              '   lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dm_typeid", oDMData("dm_typeid"))
              '   response.write "                   <input type=""button"" name=""importFromMapPointsButton"" id=""importFromMapPointsButton"" value=""Import From MapPoints"" class=""button"" onclick=""location.href='datamgr_import_from_mappoints.asp" & lcl_url_parameters & "';"" />" & vbcrlf
              'end if

              'if p_orghasfeature_customreports_datamgr then
              '   response.write "                   <input type=""button"" name=""exportButton"" id=""exportButton"" value=""Export Map-Points"" class=""button"" onclick=""openCustomReports('DATAMGR_EXPORT');"" />" & vbcrlf
              'end if

              response.write "                  </td>" & vbcrlf
              response.write "              </tr>" & vbcrlf
              response.write "            </table>" & vbcrlf
              response.write "            </p>" & vbcrlf
           'end if

           response.write "<p>" & vbcrlf
         		'response.write "<div class=""shadow"">" & vbcrlf
           response.write "<table id=""mappoints"" cellspacing=""0"" cellpadding=""2"" class=""mappoints_sortable"" style=""width:1000px"">" & vbcrlf
           response.write "  <thead>" & vbcrlf
         		response.write "  <tr>" & vbcrlf

           if p_isRootAdmin then
              response.write "      <th align=""left""><span>DM Type</span></th>" & vbcrlf
              response.write "      <th><span>DM<br />ImportID</span></th>" & vbcrlf
           end if

           response.write "      <th>&nbsp;</th>" & vbcrlf
           response.write "      <th align=""left""><span>Category</span></th>" & vbcrlf
           'response.write "      <th align=""left"">Property Address</th>" & vbcrlf
           'response.write "      <th align=""left"">Latitude</th>" & vbcrlf
           'response.write "      <th align=""left"">Longitude</th>" & vbcrlf

          'Pull all of the columns that are in the "account info" section
           lcl_display_type = "FIELDNAME"

           buildResultsList lcl_display_type, _
                            oDMData("dmid"), _
                            oDMData("dm_typeid"), _
                            oDMData("accountInfoSectionID"), _
                            lcl_row_onclick, _
                            sSCSearchField

           response.write "      <th><span>Active</span></th>" & vbcrlf
           response.write "      <th><span>Approval Status</span></th>" & vbcrlf
           response.write "      <th><span>Approved/Denied By</span></th>" & vbcrlf
           response.write "      <th>&nbsp;</th>" & vbcrlf
           response.write "  </tr>" & vbcrlf
           response.write "  </thead>" & vbcrlf

           lcl_bgcolor             = "#ffffff"
           lcl_original_categoryid = 0
        end if

       'Set up the map point color for the category
        lcl_display_mappointcolor = "&nbsp;"

        if oDMData("mappointcolor") <> "" then
           lcl_display_mappointcolor = "<img src=""mappoint_colors/bg_" & oDMData("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />"
        end if

       'BEGIN: DM Data Row ----------------------------------------------------
        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf

        if p_isRootAdmin then
           response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMData("description") & "</td>" & vbcrlf
           response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & oDMData("dm_importid") & "</td>" & vbcrlf
        end if

        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & lcl_display_mappointcolor & "</td>" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMData("categoryname")   & "</td>" & vbcrlf
        'response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """><span id=""dmdata" & oDMData("dmid") & """>" & lcl_displayAddress & "</span></td>" & vbcrlf
        'response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMData("latitude")  & "</td>" & vbcrlf
        'response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMData("longitude") & "</td>" & vbcrlf
        'Pull all of the columns that are in the "account info" section

        if oDMData("dmid") <> "" AND oDMData("dm_typeid") <> "" AND oDMData("accountInfoSectionID") <> "" then
           lcl_display_type = "FIELDVALUE"

           buildResultsList lcl_display_type, _
                            oDMData("dmid"), _
                            oDMData("dm_typeid"), _
                            oDMData("accountInfoSectionID"), _
                            lcl_row_onclick, _
                            sSCSearchField
        end if

        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """ align=""center"">" & lcl_display_active      & "</td>" & vbcrlf
'        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
'        response.write "          <span id=""approvedDeniedButtons" & iRowCount & """>" & vbcrlf

'        if lcl_show_approvedButton > 0 then
'           response.write "          <input type=""button"" name=""approveButton" & iRowCount & """ id=""approveButton" & iRowCount & """ class=""button"" value=""Approve"" onclick=""approveDenyDMID('" & iRowCount & "','A');"" />" & vbcrlf
'        end if

'        if lcl_show_deniedButton > 0 then
'           response.write "          <input type=""button"" name=""denyButton" & iRowCount & """ id=""denyButton" & iRowCount & """ class=""button"" value=""Deny"" onclick=""approveDenyDMID('" & iRowCount & "','D');"" />" & vbcrlf
'        end if

'        response.write "          </span>" & vbcrlf
'        response.write "      </td>" & vbcrlf
'        response.write "      <td align=""center"" nowrap=""nowrap""><span id=""approvedDeniedInfo" & iRowCount & """>" & lcl_display_approvedDeniedInfo & "</span></td>" & vbcrlf

        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <span id=""approvedDeniedStatus" & iRowCount & """ class=""redText"">" & lcl_approved_denied_status & "</span>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <span id=""approvedDeniedInfo" & iRowCount & """>" & lcl_approvedenied_info & "</span><br />" & vbcrlf
        response.write "          <span id=""approvedDeniedButtons" & iRowCount & """>" & vbcrlf

        if lcl_show_approvedButton > 0 then
           response.write "          <input type=""button"" name=""approveButton" & iRowCount & """ id=""approveButton" & iRowCount & """ class=""button"" value=""Approve"" onclick=""approveDenyDMID('" & iRowCount & "','A');"" />" & vbcrlf
        end if

        if lcl_show_deniedButton > 0 then
           response.write "          <input type=""button"" name=""denyButton" & iRowCount & """ id=""denyButton" & iRowCount & """ class=""button"" value=""Deny"" onclick=""approveDenyDMID('" & iRowCount & "','D');"" />" & vbcrlf
        end if

        response.write "          </span>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td class=""formlist"" align=""center"">"& vbcrlf
        response.write "          <input type=""button"" name=""delete" & iRowCount & """ id=""delete"   & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirmDelete('" & oDMData("dmid") & "');"" />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dmid" & iRowCount & """ id=""dmid" & iRowCount & """ value=""" & oDMData("dmid") & """ size=""3"" maxlength=""100"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf
       'END: DM Data Row ------------------------------------------------------

        lcl_previous_dm_typeid = oDMData("dm_typeid")

        oDMData.movenext
     loop

   		response.write "</table>" & vbcrlf
	    'response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf
     response.write "<input type=""hidden"" name=""total_waiting"" id=""total_waiting"" value=""" & lcl_waiting_count & """ size=""5"" maxlength=""10"" />" & vbcrlf
  else
     lcl_description = ""

     displayButtonRow p_isRootAdmin, _
                      p_feature, _
                      p_dm_typeid, _
                      lcl_description, _
                      lcl_url_parameters, _
                      p_orghasfeature_customreports_datamgr, _
                      lcl_url_parameters

   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No Records Available.</p>" & vbcrlf
  end if

 	oDMData.close
 	set oDMData = nothing


  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub displayButtonRow(ByVal iIsRootAdmin, ByVal iFeature, ByVal iDMTypeID, ByVal iDescription, _
                     ByVal iURLParameters, ByVal iOrgHasFeature_customReports_datamgr, _
                     ByRef lcl_url_parameters)

  sIsRootAdmin = false
  sFeature     = ""
  sDMTypeID    = 0

  if iIsRootAdmin <> "" then
     sIsRootAdmin = iIsRootAdmin
  end if

  if iFeature <> "" then
     if not containsApostrophe(iFeature) then
        sFeature = iFeature
     end if
  end if

  if iDMTypeID <> "" then
     sDMTypeID = clng(iDMTypeID)
  end if

  lcl_url_parameters = iURLParameters

  if sDMTypeID > 0 then
     if sFeature <> "" then
        lcl_addButtonLabel = getFeatureName(sFeature)
     else
        lcl_addButtonLabel = iDescription
     end if

     lcl_addButtonLabel = replace(lcl_addButtonLabel,"Maintain ","")
     lcl_addButtonURL   = "datamgr_maint.asp" & lcl_url_parameters

     response.write "            <p>" & vbcrlf
     response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "              <tr valign=""top"">" & vbcrlf
     response.write "                  <td><input type=""button"" name=""addButton"" id=""addButton"" value=""Add " & lcl_addButtonLabel & """ class=""button"" onclick=""location.href='" & lcl_addButtonURL & "'"" /></td>" & vbcrlf
     response.write "                  <td align=""right"">" & vbcrlf

     if sIsRootAdmin then
        'lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dm_typeid", sDMTypeID)
        response.write "                   <input type=""button"" name=""importFromMapPointsButton"" id=""importFromMapPointsButton"" value=""Import From MapPoints"" class=""button"" onclick=""location.href='datamgr_import_from_mappoints.asp" & lcl_url_parameters & "';"" />" & vbcrlf
        response.write "                   <input type=""button"" name=""importFromSpreadsheetButton"" id=""importFromSpreadsheetButton"" value=""Import From Spreadsheet"" class=""button"" onclick=""location.href='datamgr_import_from_spreadsheet.asp" & lcl_url_parameters & "';"" />" & vbcrlf
        response.write "                   <input type=""button"" name=""getNonValidAddressLatLongButton"" id=""getNonValidAddressLatLongButton"" value=""Retrieve Latitude/Longtitude"" class=""button"" onclick=""location.href='datamgr_get_latlong.asp" & lcl_url_parameters & "';"" />" & vbcrlf
     end if

     if p_orghasfeature_customreports_datamgr then
        response.write "                   <input type=""button"" name=""exportButton"" id=""exportButton"" value=""Export Map-Points"" class=""button"" onclick=""openCustomReports('DATAMGR_EXPORT');"" />" & vbcrlf
     end if

     response.write "                  </td>" & vbcrlf
     response.write "              </tr>" & vbcrlf
     response.write "            </table>" & vbcrlf
     response.write "            </p>" & vbcrlf
  end if

end sub

'------------------------------------------------------------------------------
sub buildResultsList(iDisplayType, iDMID, iDM_TypeID, iAccountInfoSectionID, iRowOnClick, iSCSearchField)

  dim lcl_displayType, lcl_fieldtype, lcl_fieldvalue, lcl_searchfield

  lcl_displayType = "FIELDVALUE"
  lcl_fieldtype   = ""
  lcl_fieldvalue  = ""
  lcl_searchfield = ""

  if iDisplayType <> "" then
     lcl_displayType = ucase(iDisplayType)
  end if

  if iSCSearchField <> "" then
     lcl_searchfield = ucase(iSCSearchField)
  end if

  sSQLc = "SELECT dmtf.dm_fieldid, "
  sSQLc = sSQLc & " dmtf.dm_sectionid, "
  sSQLc = sSQLc & " dmtf.section_fieldid, "
  sSQLc = sSQLc & " dmsf.fieldname, "
  sSQLc = sSQLc & " dmsf.fieldtype, "
  sSQLc = sSQLc & " dmv.dm_valueid, "
  sSQLc = sSQLc & " dmv.fieldvalue, "
  sSQLc = sSQLc & " dmtf.resultsOrder "
  sSQLc = sSQLc & "  FROM egov_dm_types_fields dmtf "
  sSQLc = sSQLc & "     LEFT OUTER JOIN egov_dm_values dmv "
  sSQLc = sSQLc & "                  ON dmtf.dm_fieldid = dmv.dm_fieldid "
  sSQLc = sSQLc & "                 AND dmv.dmid = " & iDMID
  sSQLc = sSQLc & "                 AND dmv.dm_typeid = " & iDM_TypeID
  sSQLc = sSQLc & "     LEFT OUTER JOIN egov_dm_sections_fields dmsf "
  sSQLc = sSQLc & "                  ON dmtf.section_fieldid = dmsf.section_fieldid "
  sSQLc = sSQLc & "                 AND dmsf.sectionid = " & iAccountInfoSectionID
  sSQLc = sSQLc & "  WHERE dmtf.dm_sectionid IN (SELECT dmts.dm_sectionid "
  sSQLc = sSQLc & "                              FROM egov_dm_types_sections dmts "
  sSQLc = sSQLc & "                              WHERE dmts.dm_typeid = " & iDM_TypeID
  sSQLc = sSQLc & "                                AND dmts.sectionid IN (SELECT dms.sectionid "
  sSQLc = sSQLc & "                                                       FROM egov_dm_sections dms "
  sSQLc = sSQLc & "                                                       WHERE dms.isAccountInfoSection = 1 "
  sSQLc = sSQLc & "                                                       AND dms.isActive = 1 "
  sSQLc = sSQLc & "                                                       AND dms.sectionid = " & iAccountInfoSectionID
  sSQLc = sSQLc & "                                                      ) "
  sSQLc = sSQLc & "                             ) "
  sSQLc = sSQLc & "  AND dmtf.displayInResults = 1 "

  'if lcl_displayType = "FIELDVALUE" AND lcl_sc_searchfield <> "" then
  '   lcl_fieldvalue = ucase(lcl_sc_searchfield)
  '   lcl_fieldvalue = dbsafe(lcl_fieldvalue)
  '   lcl_fieldvalue = "'%" & lcl_fieldvalue & "%'"

  '   sSQLc = sSQLc & " AND UPPER(dmv.fieldvalue) LIKE (" & lcl_fieldvalue & ") "
  'end if

  'sSQLc = sSQLc & "  ORDER BY dmtf.dm_sectionid, dmtf.dm_fieldid "
  sSQLc = sSQLc & "  ORDER BY dmtf.resultsOrder "

  set oAccountInfoColumns = Server.CreateObject("ADODB.Recordset")
  oAccountInfoColumns.Open sSQLc, Application("DSN"), 3, 1

  if not oAccountInfoColumns.eof then
     do while not oAccountInfoColumns.eof

        if lcl_displayType = "FIELDNAME" then
           response.write "      <th align=""left""><span>" & oAccountInfoColumns("fieldname") & "</span></th>" & vbcrlf
        else
           lcl_fieldtype  = oAccountInfoColumns("fieldtype")
           lcl_fieldvalue = oAccountInfoColumns("fieldvalue")

           if instr(lcl_fieldtype,"WEBSITE") > 0 OR instr(lcl_fieldtype,"EMAIL") > 0 then
              lcl_fieldvalue = buildURLDisplayValue(lcl_fieldtype, lcl_fieldvalue)
           end if

           response.write "      <td class=""formlist"" onclick=""" & iRowOnClick & """>" & lcl_fieldvalue & "</td>" & vbcrlf
        end if

        oAccountInfoColumns.movenext
     loop
  end if

  'oAccountInfoColumns.close
  set oAccountInfoColumns = nothing
end sub

'------------------------------------------------------------------------------
sub displayApprovedDeniedOptions(iSCApprovedDenied)

  lcl_sc_ad                 = "0"
  lcl_selected_viewall      = ""
  lcl_selected_viewnot      = ""
  lcl_selected_viewapproved = ""
  lcl_selected_viewdenied   = ""

  if iSCApprovedDenied <> "" then
     lcl_sc_ad = iSCApprovedDenied
  end if

  if lcl_sc_ad = "1" then
     lcl_selected_viewnot      = " selected=""selected"""
  elseif lcl_sc_ad = "2" then
     lcl_selected_viewapproved = " selected=""selected"""
  elseif lcl_sc_ad = "3" then
     lcl_selected_viewdenied   = " selected=""selected"""
  else
     lcl_selected_viewall      = " selected=""selected"""
  end if

  response.write "  <option value=""0""" & lcl_selected_viewall      & ">View all</option>" & vbcrlf
  response.write "  <option value=""1""" & lcl_selected_viewnot      & ">View all WAITING for Approval</option>" & vbcrlf
  response.write "  <option value=""2""" & lcl_selected_viewapproved & ">View all APPROVED</option>" & vbcrlf
  response.write "  <option value=""3""" & lcl_selected_viewdenied   & ">View all DENIED</option>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayDMImportIDs(iOrgID, iFeature, iSC_DMImportID)

  dim sOrgID, sFeature, sSC_DMImportID

  sOrgID         = 0
  sFeature       = ""
  sSC_DMImportID = 0

  if iOrgID <> "" then
     if isnumeric(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
  end if

  if iFeature <> "" then
     if not containsApostrophe(iFeature) then
        sFeature = iFeature
        sFeature = ucase(sFeature)
        sFeature = dbsafe(sFeature)
        sFeature = "'" & sFeature & "'"
     end if
  end if

  if iSC_DMImportID <> "" then
     sSC_DMImportID = clng(iSC_DMImportID)
  end if

  sSQL = "SELECT distinct dm_importid "
  sSQL = sSQL & " FROM egov_dm_data "
  sSQL = sSQL & " WHERE orgid = " & sOrgID
  sSQL = sSQL & " AND dm_importid <> '' "
  sSQL = sSQL & " AND dm_importid IS NOT NULL "
  sSQL = sSQL & " ORDER BY dm_importid "

  set oDisplayDMImportIDs = Server.CreateObject("ADODB.Recordset")
  oDisplayDMImportIDs.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayDMImportIDs.eof then
     do while not oDisplayDMImportIDs.eof

        if oDisplayDMImportIDs("dm_importid") = sSC_DMImportID then
           lcl_selected_dm_importid = " selected=""selected"""
        else
           lcl_selected_dm_importid = ""
        end if

        response.write "  <option value=""" & oDisplayDMImportIDs("dm_importid") & """" & lcl_selected_dm_importid & ">" & oDisplayDMImportIDs("dm_importid") & "</option>" & vbcrlf
        oDisplayDMImportIDs.movenext
     loop
  end if

  oDisplayDMImportIDs.close
  set oDisplayDMImportIDs = nothing

end sub
%>