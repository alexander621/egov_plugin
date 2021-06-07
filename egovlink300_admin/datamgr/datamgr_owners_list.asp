<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_owners_list.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists all of the owners/editors entries for in a DM Type
'
' MODIFICATION HISTORY
' 1.0 10/17/2011  David Boyer - Initial Version
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
 lcl_isRootAdmin        = False
 lcl_feature            = "datamgr_owners"
 lcl_showsearchcriteria = true

 if request("f") <> "" AND request("f") <> "datamgr_owners" then
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
    lcl_isRootAdmin = True
 end if

 lcl_featurename = getFeatureName(lcl_feature)
 lcl_dm_typeid   = getDMTypeByFeature(session("orgid"), "feature_owners", lcl_feature)

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
 'lcl_orghasfeature_customreports_datamgr = orghasfeature("customreports_datamgr")

'Check for user permissions
 lcl_userhaspermission_feature          = userhaspermission(session("userid"),lcl_feature)
 lcl_userhaspermission_feature_maintain = userhaspermission(session("userid"),lcl_feature)

'Retrieve the search options
 lcl_sc_dm_typeid      = ""
 lcl_sc_ownername      = ""
 lcl_sc_approvedDenied = ""
' lcl_sc_fromcreatedate = ""
' lcl_sc_tocreatedate   = ""
' lcl_sc_title          = ""
' lcl_sc_userid         = 0
' lcl_sc_orderby        = "createdate"

 if request("sc_dm_typeid") <> "" then
    lcl_sc_dm_typeid = request("sc_dm_typeid")
    lcl_sc_dm_typeid = clng(lcl_sc_dm_typeid)
 end if

 if request("sc_ownername") <> "" then
    lcl_sc_ownername = request("sc_ownername")
 end if

 if request("sc_approvedDenied") <> "" then
    lcl_sc_approvedDenied = request("sc_approvedDenied")
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
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_ownername",      lcl_sc_ownername)
  lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_approvedDenied", lcl_sc_approvedDenied)

  session("RedirectPage") = session("egovclientwebsiteurl") & "/admin/datamgr/datamgr_owners_list.asp" & lcl_url_parameters
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
  <link rel="stylesheet" type="text/css" href="layout_styles.css" />

<style type="text/css">
  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
 }
</style>

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>
<% '  <script type="text/javascript" src="https://github.com/jquery/jquery-ui.git"></script> %>

<script language="javascript">
<!--
function approveDenyOwnerEditor(iOwnerType, iRowID, iDMOwnerID, iAction) {
  var lcl_ownertype  = 'OWNER';
  var lcl_dm_ownerid = '';
  var lcl_action     = 'DENIED';

  if(iAction != '') {
     lcl_action = iAction;

     if(iOwnerType != '') {
        lcl_ownertype = iOwnerType;
     }

     if(iDMOwnerID != '') {
        lcl_dm_ownerid = iDMOwnerID;
     }
//alert('approveDenyOwnerEditor.asp?orgid=<%=session("orgid")%>&userid=<%=session("userid")%>&dm_ownerid=' + lcl_dm_ownerid + '&approval_action=' + lcl_action + '&isAjax=Y');
     //Approve/Deny the Owner/Editor
     $.post('approveDenyOwnerEditor.asp', {
        orgid:           '<%=session("orgid")%>',
        userid:          '<%=session("userid")%>',
        dm_ownerid:      lcl_dm_ownerid,
        approval_action: lcl_action,
        isAjax:          'Y'
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
           lcl_button_nameid = lcl_ownertype + '_' + 'denyButton' + iRowID;
           lcl_button_value  = 'Deny';
           lcl_button_action = 'DENIED';
        } else {
           lcl_status_value  = 'DENIED';
           lcl_button_nameid = lcl_ownertype + '_' + 'approveButton' + iRowID;
           lcl_button_value  = 'Approve';
           lcl_button_action = 'APPROVED';
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
        lcl_button +=   "onclick=\"approveDenyOwnerEditor('" + lcl_ownertype + "','" + iRowID + "','" + lcl_dm_ownerid + "','" + lcl_button_action + "');\" ";
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
  response.write "<form name=""datamgr_owners"" id=""datamgr_owners"" action=""datamgr_owners_list.asp"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""f"" id=""f"" size=""10"" maxlength=""50"" value=""" & lcl_feature & """ />" & vbcrlf

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
  response.write "                  <td align=""right""><span id=""screenMsg"">&nbsp;</span></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

'  if lcl_showsearchcriteria then
     lcl_total_waitingforapproval = getTotalDMID_waitingForApproval(session("orgid"), lcl_sc_dm_typeid, lcl_feature)

     lcl_dm_type_label = lcl_featurename
     lcl_dm_type_label = replace(lcl_dm_type_label,"Maintain ","")

     response.write "              <tr valign=""top"">" & vbcrlf
     response.write "                  <td colspan=""2"">" & vbcrlf
     response.write "                      <fieldset class=""fieldset"">" & vbcrlf
     response.write "                        <legend>Search Options&nbsp;</legend>" & vbcrlf
     response.write "                        <p>" & vbcrlf
     response.write "                        <table border=""0"" cellspacing=""1"" cellpadding=""0"">" & vbcrlf
     response.write "                          <tr valign=""top"">" & vbcrlf

     if lcl_showsearchcriteria then
        response.write "                           <td>" & vbcrlf
        response.write "                               DM Type:" & vbcrlf
        response.write "                           </td>" & vbcrlf
        response.write "                           <td colspan=""3"">" & vbcrlf
        response.write "                               <select name=""sc_dm_typeid"" id=""sc_dm_typeid"">" & vbcrlf
        response.write "                                 <option value=""""></option>" & vbcrlf
                                                         displayDMTypes session("orgid"), lcl_sc_dm_typeid, lcl_feature
        response.write "                               </select>" & vbcrlf
     else
        response.write "                           <td colspan=""4"">" & vbcrlf
        response.write "                               <input type=""hidden"" name=""sc_dm_typeid"" id=""sc_dm_typeid"" value=""" & lcl_sc_dm_typeid & """ />" & vbcrlf
     end if

     response.write "                              </td>" & vbcrlf
     response.write "                          </tr>" & vbcrlf
     response.write "                          <tr valign=""top"">" & vbcrlf
     response.write "                              <td>" & vbcrlf
     response.write "                                  Owner:" & vbcrlf
     response.write "                              </td>" & vbcrlf
     response.write "                              <td>" & vbcrlf
     response.write "                                  <input type=""text"" name=""sc_ownername"" id=""sc_ownername"" value=""" & lcl_sc_ownername & """ size=""30"" maxlength=""50"" />" & vbcrlf
     response.write "                              </td>" & vbcrlf
     response.write "                              <td>" & vbcrlf
     response.write "                                  Approval Status:" & vbcrlf
     response.write "                              </td>" & vbcrlf
     response.write "                              <td>" & vbcrlf
     response.write "                                  <select name=""sc_approvedDenied"" id=""sc_approvedDenied"">" & vbcrlf
                                                         displayApprovedDeniedOptions lcl_sc_approvedDenied
     response.write "                                  </select>" & vbcrlf
     response.write "                                  <div id=""waitingApprovalText"" class=""approvalStatus""></div>" & vbcrlf
     response.write "                              </td>" & vbcrlf
     response.write "                          </tr>" & vbcrlf
     response.write "                        </table>" & vbcrlf
     response.write "                        </p>" & vbcrlf
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

  '   if lcl_orghasfeature_customreports_datamgr then
  '      response.write "                  <td align=""right""><input type=""button"" name=""exportButton"" id=""exportButton"" value=""Export Map-Points"" class=""button"" onclick=""openCustomReports('DATAMGR_EXPORT');"" /></td>" & vbcrlf
  '   end if

  '   response.write "              </tr>" & vbcrlf
  '   response.write "            </table>" & vbcrlf
  '   response.write "            </p>" & vbcrlf
  'end if

                              displayOwners lcl_isRootAdmin, session("orgid"), lcl_feature, lcl_dm_typeid, _
                                            lcl_sc_dm_typeid, lcl_sc_ownername, lcl_sc_approvedDenied, lcl_url_parameters

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
sub displayOwners(p_isRootAdmin, p_orgid, p_feature, p_dm_typeid, p_sc_dm_typeid, p_sc_ownername, p_sc_approvedDenied, p_url_parameters)

 	dim iRowCount, lcl_previous_dm_typeid, lcl_waiting_count, sOrgID, sFeature
  dim lcl_sql_select, lcl_sql_from, lcl_sql_where, lcl_sql_orderby
  dim lcl_sql_approvedDenied, lcl_sql_feature, lcl_sql_dm_typeid, lcl_sql_ownername

  lcl_sql_select         = ""
  lcl_sql_from           = ""
  lcl_sql_where          = ""
  lcl_sql_orderby        = ""
  lcl_sql_approvedDenied = ""
  lcl_sql_feature        = ""
  lcl_sql_dm_typeid      = ""
  lcl_sql_ownername      = ""
  lcl_previous_dm_typeid = 0
  lcl_waiting_count      = 0
  sSC_DM_TypeID          = 0
  sSC_OwnerName          = ""
  sSC_ApprovedDenied     = 0

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  else
     sOrgID = "0"
  end if

  if p_feature <> "" then
     sFeature = ucase(p_feature)
     sFeature = dbsafe(sFeature)
     sFeature = "'" & sFeature & "'"
  else
     sFeature = ""
  end if

  if p_sc_dm_typeid <> "" then
     sSC_DM_TypeID = clng(p_sc_dm_typeid)
  end if

  if p_sc_ownername <> "" then
     sSC_OwnerName = p_sc_ownername
     sSC_OwnerName = ucase(sSC_OwnerName)
     sSC_OwnerName = dbsafe(sSC_OwnerName)
  end if

  if p_sc_approveddenied <> "" then
     sSC_ApprovedDenied = clng(p_sc_approveddenied)
  end if

 'Setup SELECT
  lcl_sql_select = " SELECT "
  lcl_sql_select = lcl_sql_select & " dmo.dm_ownerid, "
  lcl_sql_select = lcl_sql_select & " dmo.dmid, "
  lcl_sql_select = lcl_sql_select & " dm.dm_typeid, "
  lcl_sql_select = lcl_sql_select & " dmt.description, "
  lcl_sql_select = lcl_sql_select & " dmt.accountInfoSectionID, "
  lcl_sql_select = lcl_sql_select & " dm.categoryid, "
  lcl_sql_select = lcl_sql_select & " dmc.categoryname, "
  lcl_sql_select = lcl_sql_select & " dmo.userid, "
  lcl_sql_select = lcl_sql_select & " u.userfname + ' ' + u.userlname AS ownername, "
  lcl_sql_select = lcl_sql_select & " dmo.ownertype, "
  lcl_sql_select = lcl_sql_select & " dmo.isApprovedDeniedByAdmin, "
  lcl_sql_select = lcl_sql_select & " dmo.isApproved, "
  lcl_sql_select = lcl_sql_select & " dmo.approvedeniedbyid, "
  lcl_sql_select = lcl_sql_select & " dmo.approvedeniedbydate, "
  lcl_sql_select = lcl_sql_select & " CASE "
'  lcl_sql_select = lcl_sql_select &      " WHEN dmo.isApproved = 0 AND (dmo.approvedeniedbydate = '' OR dmo.approvedeniedbydate IS NULL) THEN "
'  lcl_sql_select = lcl_sql_select &           " '' "
  lcl_sql_select = lcl_sql_select &      " WHEN dmo.isApprovedDeniedByAdmin = 1 THEN "
  lcl_sql_select = lcl_sql_select &           " u3.firstname + ' ' + u3.lastname "
  lcl_sql_select = lcl_sql_select &      " ELSE "
  lcl_sql_select = lcl_sql_select &           " u2.userfname + ' ' + u2.userlname "
  lcl_sql_select = lcl_sql_select & " END AS approvedeniedbyname "

 'Setup FROM
  lcl_sql_from = " FROM egov_dm_owners as dmo "
  lcl_sql_from = lcl_sql_from & " INNER JOIN egov_dm_data AS dm ON dmo.dmid = dm.dmid AND dm.orgid = "                & sOrgID
  lcl_sql_from = lcl_sql_from & " INNER JOIN egov_dm_types AS dmt ON dm.dm_typeid = dmt.dm_typeid AND dmt.orgid = "   & sOrgID
  lcl_sql_from = lcl_sql_from & " LEFT OUTER JOIN egov_users u ON dmo.userid = u.userid AND u.orgid = "               & sOrgID
  lcl_sql_from = lcl_sql_from & " LEFT OUTER JOIN egov_users u2 ON dmo.approvedeniedbyid = u2.userid AND u2.orgid = " & sOrgID
  lcl_sql_from = lcl_sql_from & " LEFT OUTER JOIN users u3 ON dmo.approvedeniedbyid = u3.userid AND u3.orgid = "      & sOrgID
  lcl_sql_from = lcl_sql_from & " LEFT OUTER JOIN egov_dm_categories dmc ON dm.categoryid = dmc.categoryid "

 'Setup WHERE
  lcl_sql_where = " WHERE dmo.orgid = " & sOrgID
  lcl_sql_where = lcl_sql_where & " AND NOT (dm.isApproved = 0 AND dm.approvedeniedbydate <> '' AND dm.approvedeniedbydate IS NOT NULL) "

 'Search Options / Feature
  if p_feature <> "" AND p_feature <> "datamgr_maint" AND p_feature <> "datamgr_owners" then
       lcl_sql_feature = " AND UPPER(dmt.feature_owners) = " & sFeature
  end if

  if p_feature <> "" AND p_feature <> "datamgr_maint" AND p_feature <> "datamgr_owners" then
     lcl_sql_dm_typeid = " AND dm.dm_typeid = " & p_dm_typeid
  else
    'Setup the WHERE clause with the search option values.
     if sSC_DM_TypeID > 0 then
        lcl_sql_dm_typeid = " AND dm.dm_typeid = " & sSC_DM_TypeID
     end if
  end if

  if sSC_OwnerName <> "" then
     lcl_sql_ownername = " AND UPPER(u.userfname) + ' ' + UPPER(u.userlname) LIKE ('%" & sSC_OwnerName & "%') "
  end if

 'Setup ORDER BY
  lcl_sql_orderby = " ORDER BY 9, dmt.description "

 'Determine if we are showing approved and/or denied DM Data
 '0 = View All
 '1 = View all WAITING for approval
 '2 = View all APPROVED
 '3 = View all DENIED
'  if sSC_ApprovedDenied <> "" then
     if sSC_ApprovedDenied > 0 then
        if sSC_ApprovedDenied = 1 then
           lcl_sql_approvedDenied = "= 0 AND (dmo.approvedeniedbydate = '' OR dmo.approvedeniedbydate IS NULL) "
        elseif sSC_ApprovedDenied = 2 then
           lcl_sql_approvedDenied = "= 1 "
        elseif sSC_ApprovedDenied = 3 then
           lcl_sql_approvedDenied = "= 0 AND dmo.approvedeniedbydate <> '' "
        end if

        lcl_sql_approvedDenied = " AND dmo.isApproved " & lcl_sql_approvedDenied

     end if
'  end if

 'BEGIN: Owners - Approved ----------------------------------------------------
  sSQL = lcl_sql_select
  sSQL = sSQL & lcl_sql_from
  sSQL = sSQL & lcl_sql_where
  sSQL = sSQL & " AND dmo.isApproved = 1 "
  sSQL = sSQL & lcl_sql_feature
  sSQL = sSQL & lcl_sql_dm_typeid
  sSQL = sSQL & lcl_sql_ownername
  sSQL = sSQL & lcl_sql_approvedDenied
 'END: Owners - Approved ------------------------------------------------------

 'BEGIN: Owners - Denied ------------------------------------------------------
  sSQL = sSQL & " UNION ALL "
  sSQL = sSQL & lcl_sql_select
  sSQL = sSQL & lcl_sql_from
  sSQL = sSQL & lcl_sql_where
  sSQL = sSQL & " AND dmo.isApproved = 0 "
  sSQL = sSQL & " AND dmo.approvedeniedbydate <> '' "
  sSQL = sSQL & " AND dmo.approvedeniedbydate IS NOT NULL "
  sSQL = sSQL & lcl_sql_feature
  sSQL = sSQL & lcl_sql_dm_typeid
  sSQL = sSQL & lcl_sql_ownername
  sSQL = sSQL & lcl_sql_approvedDenied
 'END: Owners - Denied --------------------------------------------------------

 'BEGIN: Owners - Waiting for Approval ----------------------------------------
  sSQL = sSQL & " UNION ALL "
  sSQL = sSQL & lcl_sql_select
  sSQL = sSQL & lcl_sql_from
  sSQL = sSQL & lcl_sql_where
  sSQL = sSQL & " AND dmo.isApproved = 0 "
  sSQL = sSQL & " AND (dmo.approvedeniedbydate = '' OR dmo.approvedeniedbydate IS NULL) "
  sSQL = sSQL & lcl_sql_feature
  sSQL = sSQL & lcl_sql_dm_typeid
  sSQL = sSQL & lcl_sql_ownername
  sSQL = sSQL & lcl_sql_approvedDenied
 'END: Owners - Waiting for Approval ------------------------------------------

 'BEGIN: ORDER BY -------------------------------------------------------------
  sSQL = sSQL & lcl_sql_orderby
 'END: ORDER BY ---------------------------------------------------------------

  session("CR_DATAMGR_EXPORT") = sSQL

 	set oDMOwners = Server.CreateObject("ADODB.Recordset")
	 oDMOwners.Open sSQL, Application("DSN"), 3, 1
	
 	if not oDMOwners.eof then
     do while not oDMOwners.eof
        lcl_bgcolor     = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_row_onclick = ""
     			iRowCount       = iRowCount + 1

       'Setup the onclick
        'lcl_feature_maint = getFeatureFromDMType(oDMOwners("dm_typeid"), "feature_maintain")
        'lcl_row_onclick   = setupUrlParameters("", "f", lcl_feature_maint)
        'lcl_row_onclick   = setupUrlParameters(lcl_row_onclick, "dmid", oDMOwners("dmid"))
        'lcl_row_onclick   = "location.href='datamgr_maint.asp" & lcl_row_onclick & "';"

       'Build the "active" display value
'        lcl_display_active = "&nbsp;"

'        if oDMOwners("isActive") then
'           lcl_display_active = "Y"
'        end if

       'Set up Approve/Deny Buttons for display
        lcl_show_approvedButton        = 1
        lcl_show_deniedButton          = 1
        lcl_display_approvedDeniedInfo = ""
        lcl_approved_denied_status     = ""
        lcl_approvedeniedbyname        = oDMOwners("approvedeniedbyname")

        'if oDMOwners("isApprovedDeniedByAdmin") then
        '   lcl_approvedeniedbyname = oDMOwners("approvedeniedbyname_admin")
        'else
        '   lcl_approvedeniedbyname = oDMOwners("approvedeniedbyname_public")
        'end if

        lcl_approvedenied_info = formatAdminActionsInfo(lcl_approvedeniedbyname, oDMOwners("approvedeniedbydate"))

        'if lcl_approvedenied_info <> "" then
        if oDMOwners("isApproved") then
           lcl_show_approvedButton    = 0
           lcl_approved_denied_status = "APPROVED"
        else
           if oDMOwners("approvedeniedbydate") <> "" then
              lcl_show_deniedButton      = 0
              lcl_approved_denied_status = "DENIED"
           else
              lcl_approved_denied_status = "WAITING FOR<br />APPROVAL"
              lcl_waiting_count          = lcl_waiting_count + 1
           end if
        end if

        lcl_display_approvedDeniedInfo = "<span class=""redText"">" & lcl_approved_denied_status & "</span><br />"

        if lcl_waiting_count = 0 then
           lcl_display_approvedDeniedInfo = lcl_display_approvedDeniedInfo & lcl_approvedenied_info
        end if

       'If the DM_TypeID is NOT equal to the previous DM_TypeID in the loop then close the table and open a new one.
       'The reason for this is because each DM Type has it's own "account info" fields and the columns would never line up in the results list.
       'This is mainly just for the ROOT ADMIN as clients will never see the "DM Type" column and/or search option.
        if oDMOwners("dm_typeid") <> lcl_previous_dm_typeid then
           if iRowCount > 1 then
            		response.write "</table>" & vbcrlf
         	    response.write "</div>" & vbcrlf
              response.write "</p>" & vbcrlf
           end if

           response.write "<p>" & vbcrlf
         		response.write "<div class=""shadow"">" & vbcrlf
           response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"">" & vbcrlf
         		response.write "  <tr>" & vbcrlf
           response.write "      <th align=""left"">Owner</th>" & vbcrlf
           response.write "      <th>Approval Status</th>" & vbcrlf
           response.write "      <th>Approved/Denied By</th>" & vbcrlf

           if p_isRootAdmin then
              response.write "      <th align=""left"">DM Type</th>" & vbcrlf
           end if

           response.write "      <th align=""left"">Category</th>" & vbcrlf

          'Pull all of the columns that are in the "account info" section
           lcl_display_type = "FIELDNAME"

           buildResultsList lcl_display_type, oDMOwners("dmid"), oDMOwners("dm_typeid"), oDMOwners("accountInfoSectionID"), lcl_row_onclick

           response.write "      <th>&nbsp;</th>" & vbcrlf
           response.write "  </tr>" & vbcrlf

           lcl_bgcolor             = "#ffffff"
           lcl_original_categoryid = 0
        end if

       'Set up the map point color for the category
        lcl_display_mappointcolor = "&nbsp;"

'        if oDMOwners("mappointcolor") <> "" then
'           lcl_display_mappointcolor = "<img src=""mappoint_colors/bg_" & oDMOwners("mappointcolor") & ".jpg"" width=""15"" height=""10"" style=""border:1pt solid #000000"" valign=""middle"" />"
'        end if

       'BEGIN: DM Data Row ----------------------------------------------------
        lcl_owner_url       = session("egovclientwebsiteurl") & "/admin/dirs/update_citizen.asp?userid=" & oDMOwners("userid")
        lcl_owner_mouseover = " onMouseOver=""tooltip.show('Click to Edit User');"""
        lcl_owner_mouseout  = " onMouseOut=""tooltip.hide();"""

        response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & vbcrlf
        response.write "          <input type=""hidden"" name=""dmid" & iRowCount & """ id=""dmid" & iRowCount & """ value=""" & oDMOwners("dmid") & """ size=""3"" maxlength=""100"" />" & vbcrlf
        response.write "          <input type=""hidden"" name=""dm_ownerid" & iRowCount & """ id=""dm_ownerid" & iRowCount & """ value=""" & oDMOwners("dm_ownerid") & """ size=""3"" maxlength=""100"" />" & vbcrlf
        response.write "          <a href=""" & lcl_owner_url & """" & lcl_owner_mouseover & lcl_owner_mouseout & ">" & oDMOwners("ownername") & "</a>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <span id=""approvedDeniedStatus" & iRowCount & """ class=""redText"">" & lcl_approved_denied_status & "</span>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <span id=""approvedDeniedInfo" & iRowCount & """>" & lcl_approvedenied_info & "</span><br />" & vbcrlf
        response.write "          <span id=""approvedDeniedButtons" & iRowCount & """>" & vbcrlf

        if lcl_show_approvedButton > 0 then
           response.write "          <input type=""button"" name=""approveButton" & iRowCount & """ id=""approveButton" & iRowCount & """ class=""button"" value=""Approve"" onclick=""approveDenyOwnerEditor('" & oDMOwners("ownertype") & "','" & iRowCount & "','" & oDMOwners("dm_ownerid") & "','APPROVED');"" />" & vbcrlf
           'response.write "          <input type=""button"" name=""approveButton" & iRowCount & """ id=""approveButton" & iRowCount & """ class=""button"" value=""Approve"" onclick=""approveDenyOwner('" & iRowCount & "','A');"" />" & vbcrlf
        end if

        if lcl_show_deniedButton > 0 then
           response.write "          <input type=""button"" name=""denyButton" & iRowCount & """ id=""denyButton" & iRowCount & """ class=""button"" value=""Deny"" onclick=""approveDenyOwnerEditor('" & oDMOwners("ownertype") & "','" & iRowCount & "','" & oDMOwners("dm_ownerid") & "','DENIED');"" />" & vbcrlf
           'response.write "          <input type=""button"" name=""denyButton" & iRowCount & """ id=""denyButton" & iRowCount & """ class=""button"" value=""Deny"" onclick=""approveDenyOwner('" & iRowCount & "','D');"" />" & vbcrlf
        end if

        response.write "          </span>" & vbcrlf
        response.write "      </td>" & vbcrlf

        if p_isRootAdmin then
           response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMOwners("description") & "</td>" & vbcrlf
        end if

        response.write "      <td class=""formlist"" onclick=""" & lcl_row_onclick & """>" & oDMOwners("categoryname") & "</td>" & vbcrlf

        'Pull all of the columns that are in the "account info" section
        if oDMOwners("dmid") <> "" AND oDMOwners("dm_typeid") <> "" AND oDMOwners("accountInfoSectionID") <> "" then
           lcl_display_type = "FIELDVALUE"

           buildResultsList lcl_display_type, oDMOwners("dmid"), oDMOwners("dm_typeid"), oDMOwners("accountInfoSectionID"), lcl_row_onclick
        end if

        lcl_dm_url = session("egovclientwebsiteurl") & "/admin/datamgr/datamgr_maint.asp"
        lcl_dm_url = lcl_dm_url & "?f="    & lcl_feature
        lcl_dm_url = lcl_dm_url & "&dmid=" & oDMOwners("dmid")

        response.write "      <td class=""formlist"" align=""center""><input type=""button"" name=""editDM" & iRowCount & """ id=""editDM" & iRowCount & """ value=""Edit " & oDMOwners("description") & """ class=""button"" onclick=""location.href='" & lcl_dm_url & "';"" /></td>" & vbcrlf
        response.write "  </tr>"  & vbcrlf
       'END: DM Data Row ------------------------------------------------------

        lcl_previous_dm_typeid = oDMOwners("dm_typeid")

        oDMOwners.movenext
     loop

   		response.write "</table>" & vbcrlf
	    response.write "</div>" & vbcrlf
     response.write "</p>" & vbcrlf
     response.write "<input type=""hidden"" name=""total_waiting"" id=""total_waiting"" value=""" & lcl_waiting_count & """ size=""5"" maxlength=""10"" />" & vbcrlf
  else
   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No Records Available.</p>" & vbcrlf
  end if

 	oDMOwners.close
 	set oDMOwners = nothing

end sub

'------------------------------------------------------------------------------
sub buildResultsList(iDisplayType, iDMID, iDM_TypeID, iAccountInfoSectionID, iRowOnClick)

  lcl_displayType = "FIELDVALUE"

  if iDisplayType <> "" then
     lcl_displayType = ucase(iDisplayType)
  end if

  sSQLc = "  SELECT dmtf.dm_fieldid, "
  sSQLc = sSQLc & "  dmtf.dm_sectionid, "
  sSQLc = sSQLc & "  dmtf.section_fieldid, "
  sSQLc = sSQLc & "  dmsf.fieldname, "
  sSQLc = sSQLc & "  dmsf.fieldtype, "
  sSQLc = sSQLc & "  dmv.dm_valueid, "
  sSQLc = sSQLc & "  dmv.fieldvalue "
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
  sSQLc = sSQLc & "  ORDER BY dmtf.dm_sectionid, dmtf.dm_fieldid "

  set oAccountInfoColumns = Server.CreateObject("ADODB.Recordset")
  oAccountInfoColumns.Open sSQLc, Application("DSN"), 3, 1

  if not oAccountInfoColumns.eof then
     do while not oAccountInfoColumns.eof

        if lcl_displayType = "FIELDNAME" then
           response.write "      <th align=""left"">" & oAccountInfoColumns("fieldname") & "</th>" & vbcrlf
        else
           response.write "      <td class=""formlist"" onclick=""" & iRowOnClick & """>" & oAccountInfoColumns("fieldvalue") & "</td>" & vbcrlf
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
function getTotalDMID_waitingForApproval(iOrgID, iSC_DM_TypeID, iFeature)
   lcl_return = 0

   

   getTotalDMID_waitingForApproval = lcl_return

end function
%>