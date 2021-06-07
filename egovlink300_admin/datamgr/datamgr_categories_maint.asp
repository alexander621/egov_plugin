<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_categories_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a DM Category
'
' MODIFICATION HISTORY
' 1.0 04/26/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

'Retrieve the categoryid to be maintained.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 if request("categoryid") <> "" then
    lcl_categoryid = request("categoryid")

    if isnumeric(lcl_categoryid) then
       lcl_screen_mode = "EDIT"
       lcl_sendToLabel = "Update"
    end if
 else
    lcl_screen_mode = "ADD"
    lcl_sendToLabel = "Create"
    lcl_categoryid  = 0
 end if

'Determine if the user has access to maintain
'Also determine how the user is accessing the screen.
 lcl_feature     = "datamgr_types_maint"
 lcl_featurename = ""
 lcl_dm_typeid   = 0

 if request("f") <> "" then
    lcl_feature = request("f")
 end if

'Retrieve the DM_TypeID
 if request("dm_typeid") <> "" then
    lcl_dm_typeid = request("dm_typeid")
 else
    lcl_dm_typeid = getDMTypeByFeature(session("orgid"), "feature_maintain_fields", lcl_feature)

    if lcl_dm_typeid = 0 then
      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 lcl_dm_typeid = clng(lcl_dm_typeid)
 lcl_pagetitle = getFeatureName(lcl_feature)
 lcl_pagetitle = lcl_pagetitle & " [Maintain Categories]"
 lcl_success   = request("success")

'Retrieve the search options
 lcl_sc_categoryname = ""

 if request("sc_categoryname") <> "" then
    lcl_sc_categoryname = request("sc_categoryname")
 end if

'Build return parameters
 lcl_url_parameters = ""
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dm_typeid",    lcl_dm_typeid)
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "sc_categoryname", lcl_sc_categoryname)

 if lcl_feature <> "datamgr_types_maint" then
    lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "f", lcl_feature)
 end if

'Check for org features
 lcl_orghasfeature_feature          = orghasfeature(lcl_feature)
 lcl_orghasfeature_feature_maintain = orghasfeature(lcl_feature)

'Check for user permissions
 lcl_userhaspermission_feature          = userhaspermission(session("userid"),lcl_feature)
 lcl_userhaspermission_feature_maintain = userhaspermission(session("userid"),lcl_feature)

'Set up local variables
 lcl_orgid               = session("orgid")
 lcl_categoryname        = ""
 lcl_isActive            = 1
 lcl_createdbyid         = 0
 lcl_createdbydate       = ""
 lcl_createdbyname       = ""
 lcl_lastmodifiedbyid    = 0
 lcl_lastmodifiedbydate  = ""
 lcl_lastmodifiedbyname  = ""
 lcl_parent_categoryid   = 0
 lcl_isApproved          = 1
 lcl_approvedeniedbyid   = 0
 lcl_approvedeniedbydate = ""
 lcl_mappointcolor       = ""
 lcl_checked_isactive    = " checked=""checked"""

 if lcl_screen_mode = "EDIT" then
   'Retrieve all of the data for the DM Category
    sSQL = "SELECT dmc.categoryid, "
    sSQL = sSQL & " dmc.categoryname, "
    sSQL = sSQL & " dmc.orgid, "
    sSQL = sSQL & " dmc.dm_typeid, "
    sSQL = sSQL & " dmc.isActive, "
    sSQL = sSQL & " dmc.createdbyid, "
    sSQL = sSQL & " dmc.createdbydate, "
    sSQL = sSQL & " dmc.lastmodifiedbyid, "
    sSQL = sSQL & " dmc.lastmodifiedbydate, "
    sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
    sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname, "
    sSQL = sSQL & " dmc.parent_categoryid, "
    sSQL = sSQL & " dmc.isApproved, "
    sSQL = sSQL & " dmc.approvedeniedbyid, "
    sSQL = sSQL & " dmc.approvedeniedbydate, "
    sSQL = sSQL & " dmc.mappointcolor "
    sSQL = sSQL & " FROM egov_dm_categories dmc "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON dmc.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON dmc.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
    sSQL = sSQL & " WHERE dmc.categoryid = " & lcl_categoryid

    set oDMCategory = Server.CreateObject("ADODB.Recordset")
    oDMCategory.Open sSQL, Application("DSN"), 3, 1

    if not oDMCategory.eof then
       lcl_categoryid          = oDMCategory("categoryid")
       lcl_categoryname        = oDMCategory("categoryname")
       lcl_orgid               = oDMCategory("orgid")
       lcl_dm_typeid           = oDMCategory("dm_typeid")
       lcl_isActive            = oDMCategory("isActive")
       lcl_createdbyid         = oDMCategory("createdbyid")
       lcl_createdbydate       = oDMCategory("createdbydate")
       lcl_createdbyname       = oDMCategory("createdbyname")
       lcl_lastmodifiedbyid    = oDMCategory("lastmodifiedbyid")
       lcl_lastmodifiedbydate  = oDMCategory("lastmodifiedbydate")
       lcl_lastmodifiedbyname  = oDMCategory("lastmodifiedbyname")
       lcl_parent_categoryid   = oDMCategory("parent_categoryid")
       lcl_isApproved          = oDMCategory("isApproved")
       lcl_approvedeniedbyid   = oDMCategory("approvedeniedbyid")
       lcl_approvedeniedbydate = oDMCategory("approvedeniedbydate")
       lcl_mappointcolor       = oDMCategory("mappointcolor")

      'Determine if the checkbox(es) are checked or not
       if not oDMCategory("isActive") then
          lcl_checked_isactive = ""
       end if
    else

       lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "success", "NE")

       response.redirect("datamgr_categories_list.asp" & lcl_url_parameters)
    end if

    oDMCategory.close
    set oDMCategory = nothing
 end if

'Get the description for the DM TypeID 
 lcl_displayDMT_description = getDMTypeDescription(lcl_dm_typeid)

'Format the created/last modified by info
 lcl_displayCreatedByInfo      = setupUserMaintLogInfo(lcl_createdbyname, lcl_createdbydate)
 lcl_displayLastModifiedByInfo = setupUserMaintLogInfo(lcl_lastmodifiedbyname, lcl_lastmodifiedbydate)

'Check to see if this category has been associated to a DM Type to determine if the category can been deleted.
 lcl_categoryExistsOnDMType = checkForDefaultCategoryOnDMTypes(lcl_categoryid)

 if lcl_categoryExistsOnDMType then
    lcl_canDelete = False
 else
    lcl_canDelete = True
 end if

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = lcl_onload & "setMaxLength();"

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"
    lcl_onload = lcl_onload & "window.opener.location.reload();"
 end if

 dim lcl_scripts
%>
<html>
<head>
  <title>E-Gov Administration Console {<%=lcl_pagetitle%> - <%=lcl_screen_mode%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="layout_styles.css" />

 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>
<% '  <script type="text/javascript" src="https://github.com/jquery/jquery-ui.git"></script> %>

<script language="javascript">
$(document).ready(function() {

  $('#sub_searchButton').click(function() {
    var lcl_searchvalue = $('#sub_sc_categoryname').val();
    var lcl_foundCount  = 0;

    //Hide all of the rows
    $('.subCategoryRow').each(function() {

      //Get the "id" for the current <TR> in the loop
      //var lcl_rowid = $(this).attr("id");
      var lcl_rowid = $(this).prop("id");

      //Once we have the "row id" we need only the row number
      var lcl_id = lcl_rowid.replace("subcategoryrow","");

      //Get the categoryname so we can perform the search
      var lcl_value   = $('#sub_categoryname' + lcl_id).val();
      var lcl_showRow = false;

      //Compare the value in the current row in the loop to the search value
      if(lcl_searchvalue == '') {
         lcl_showRow = true;
      } else {
         lcl_value       = lcl_value.toUpperCase();
         lcl_searchvalue = lcl_searchvalue.toUpperCase();

         if(lcl_value.indexOf(lcl_searchvalue) > -1) {
            lcl_showRow = true;
         }
      }

      //Determine if we are showing/hiding the current row.
      if(lcl_showRow) {
         lcl_foundCount = lcl_foundCount + 1;
         $('#' + lcl_rowid).show("slow");
      } else {
         $('#' + lcl_rowid).hide("slow");
      }
    });
  });

  $('#sub_addSubCategoryButton').click(function() {
    //$('#sub_sc_categoryname').attr('disabled','disabled');
    //$('#sub_searchButton').attr('disabled','disabled');
    $('#sub_sc_categoryname').prop('disabled','disabled');
    $('#sub_searchButton').prop('disabled','disabled');

    var lcl_total   = $('#totalsubcategories').val();
    //var lcl_bgcolor = $('#subcategoryrow' + lcl_total).attr('bgcolor');
    var lcl_bgcolor = $('#subcategoryrow' + lcl_total).prop('bgcolor');

    //Determine what the rowid and total row count is
    var num           = new Number(lcl_total);
    var lcl_new_total = (num + 1);
    var lcl_new_rowid = lcl_new_total.toString();
    var lcl_row_html  = "";

    if(lcl_bgcolor == "#eeeeee") {
       lcl_bgcolor = "#ffffff";
    } else {
       lcl_bgcolor = "#eeeeee";
    }

    //Build the new row
    lcl_row_html += '  <tr id="subcategoryrow' + lcl_new_rowid + '" class=""subCategoryRow"" align=""center"" bgcolor="' + lcl_bgcolor + '" valign="top">';
    lcl_row_html += '      <td align="left" nowrap="nowrap">';
    lcl_row_html += '          <input type="hidden" name="sub_categoryid' + lcl_new_rowid + '" id="sub_categoryid' + lcl_new_rowid + '" value="0" size="3" maxlength="100" />';
    lcl_row_html += '          <input type="text" name="sub_categoryname' + lcl_new_rowid + '" id="sub_categoryname' + lcl_new_rowid + '" value="" size="30" maxlength="100" onchange="clearMsg(\'sub_categoryname' + lcl_new_rowid + '\');" />';
    lcl_row_html += '      </td>';
    lcl_row_html += '      <td align="center">';
    lcl_row_html += '          <input type="checkbox" name="sub_delete' + lcl_new_rowid + '" id="sub_delete' + lcl_new_rowid + '" value="Y" />';
    lcl_row_html += '      </td>';
    lcl_row_html += '      <td colspan="3">&nbsp;</td>';
    lcl_row_html += '  </tr>';

    //Append the new row to the table and increment the sub-catgories total.
    $('#subcategories_table').append(lcl_row_html);
    $('#totalsubcategories').val(lcl_new_rowid);
  });
});

function approveDenyCategory(iRowID, iAction) {
  var lcl_categoryid = $('#sub_categoryid' + iRowID).val();
  var lcl_isApproved = false;

  if(iAction != '') {
     if(iAction == 'A') {
        lcl_isApproved = true;
     }
//alert('approveDenyCategory.asp?userid=<%=session("userid")%>&orgid=<%=session("orgid")%>&categoryid=' + lcl_categoryid + '&isApproved=' + lcl_isApproved + '&isAjax=Y');
     //Build the sub-category list
     $.post('approveDenyCategory.asp', {
        userid:           '<%=session("userid")%>',
        orgid:            '<%=session("orgid")%>',
        categoryid:       lcl_categoryid,
        isApproved:       lcl_isApproved,
        isAjax:           'Y'
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
           lcl_button_nameid = 'sub_denyButton' + iRowID;
           lcl_button_value  = 'Deny';
           lcl_button_action = 'D';
        } else {
           lcl_status_value  = 'DENIED';
           lcl_button_nameid = 'sub_approveButton' + iRowID;
           lcl_button_value  = 'Approve';
           lcl_button_action = 'A';
        }

        //Build the approve/deny info
        lcl_display_status = '<span class="redText">' + lcl_status_value + '</span><br />';

        //Build the approve/deny button
        lcl_button  = "<input ";
        lcl_button +=   "type='button' ";
        lcl_button +=   "class='button' ";
        lcl_button +=   "name='"  + lcl_button_nameid + "' ";
        lcl_button +=   "id='"    + lcl_button_nameid + "' ";
        lcl_button +=   "value='" + lcl_button_value  + "' ";
        lcl_button +=   "onclick='approveDenyCategory(\"" + iRowID + "\",\"" + lcl_button_action + "\");' ";
        lcl_button += "/>";

        $('#approvedDeniedStatus'  + iRowID).html(lcl_display_status);
        $('#approvedDeniedInfo'    + iRowID).html(lcl_display_info);
        $('#approvedDeniedButtons' + iRowID).html(lcl_button);
     });

  }
}

function deleteCategory(iRowID) {
  //Determine if we are enabling/disabling the row
  //var lcl_checked_delete   = $('#sub_delete' + iRowID).attr('checked');
  var lcl_checked_delete   = $('#sub_delete' + iRowID).prop('checked');
  var lcl_enabled_disabled = '';

  if(lcl_checked_delete) {
     lcl_enabled_disabled = 'disabled';
  }

  //$('#sub_categoryname'  + iRowID).attr('disabled',lcl_enabled_disabled);
  //$('#mergeIntoCategory' + iRowID).attr('disabled',lcl_enabled_disabled);
  $('#sub_categoryname'  + iRowID).prop('disabled',lcl_enabled_disabled);
  $('#mergeIntoCategory' + iRowID).prop('disabled',lcl_enabled_disabled);


  if($('#sub_approveButton' + iRowID)) {
     //$('#sub_approveButton' + iRowID).attr('disabled',lcl_enabled_disabled);
     $('#sub_approveButton' + iRowID).prop('disabled',lcl_enabled_disabled);
  }

  if($('#sub_denyButton' + iRowID)) {
     //$('#sub_denyButton' + iRowID).attr('disabled',lcl_enabled_disabled);
     $('#sub_denyButton' + iRowID).prop('disabled',lcl_enabled_disabled);
  }
}

function mergeCategory(iRowID) {
  lcl_currentrow_value      = $('#sub_categoryid'    + iRowID).val();
  lcl_mergerow_value        = $('#mergeIntoCategory' + iRowID).val();
  lcl_isSubCategory         = true;
  lcl_currentrow_categoryid = 0
  lcl_mergerow_categoryid   = 0
  lcl_enableFields          = true;
  lcl_return_false_count    = 0;

  //Clear any existing error message for this row
  clearMsg('mergeIntoCategory' + iRowID);

  //Disable all fields initially
  //$('#sub_categoryname' + iRowID).attr('disabled','disabled');
  //$('#sub_delete'       + iRowID).attr('disabled','disabled');
  $('#sub_categoryname' + iRowID).prop('disabled','disabled');
  $('#sub_delete'       + iRowID).prop('disabled','disabled');

  if($('#sub_approveButton' + iRowID)) {
     //$('#sub_approveButton' + iRowID).attr('disabled','disabled');
     $('#sub_approveButton' + iRowID).prop('disabled','disabled');
  }

  if($('#sub_denyButton' + iRowID)) {
     //$('#sub_denyButton' + iRowID).attr('disabled','disabled');
     $('#sub_denyButton' + iRowID).prop('disabled','disabled');
  }

  if(lcl_mergerow_value.indexOf('PC') > -1) {
     lcl_isSubCategory = false;
  }

  if(lcl_isSubCategory) {
     lcl_currentrow_categoryid = lcl_currentrow_value.replace('SC','');
     lcl_mergerow_categoryid   = lcl_mergerow_value.replace('SC','');

     if(lcl_currentrow_categoryid == lcl_mergerow_categoryid) {
        $('#mergeIntoCategory').val('');
        inlineMsg(document.getElementById('mergeIntoCategory' + iRowID).id,'<strong>Invalid Value: </strong>Cannot select the same sub-category to merge with.',10,'mergeIntoCategory' + iRowID);
        $('#mergeIntoCategory' + iRowID).focus();
        lcl_return_false_count = lcl_return_false_count + 1;
     } else {
        //alert(lcl_mergerow_categoryid);
        if(lcl_mergerow_categoryid == '') {
           lcl_enableFields = true;
        } else {
           lcl_enableFields = false;
        }
     }

  } else {
     $('#mergeIntoCategory' + iRowID).val('');
     inlineMsg(document.getElementById('mergeIntoCategory' + iRowID).id,'<strong>Invalid Value: </strong>Cannot select a parent category to merge with.',10,'mergeIntoCategory' + iRowID);
     $('#mergeIntoCategory' + iRowID).focus();
     lcl_return_false_count = lcl_return_false_count + 1;
  }

  //Determine if we need to enable fields
  if(lcl_enableFields) {

     //Enable all of the row fields, if needed.
     //$('#sub_categoryname' + iRowID).attr('disabled','');
     //$('#sub_delete'       + iRowID).attr('disabled','');
     $('#sub_categoryname' + iRowID).prop('disabled','');
     $('#sub_delete'       + iRowID).prop('disabled','');

     if($('#sub_approveButton' + iRowID)) {
        //$('#sub_approveButton' + iRowID).attr('disabled','');
        $('#sub_approveButton' + iRowID).prop('disabled','');
     }

     if($('#sub_denyButton' + iRowID)) {
        //$('#sub_denyButton' + iRowID).attr('disabled','');
        $('#sub_denyButton' + iRowID).prop('disabled','');
     }
  }

  return lcl_return_false_count;

}

function saveDMChanges() {
  clearScreenMsg();
  $('#user_action').val('UPDATE');
  //$('#datamgr_types_maint').attr('action','datamgr_action.asp');
  $('#datamgr_types_maint').prop('action','datamgr_action.asp');
  $('#datamgr_types_maint').submit();
}

function confirmDelete() {
  lcl_cname = document.getElementById("categoryname").value;

  var r = confirm("Are you sure you want to delete this category: '" + lcl_cname + "'?");
  if (r==true) {

    <%
      lcl_delete_params = lcl_url_parameters
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "user_action", "DELETE")
      lcl_delete_params = setupUrlParameters(lcl_delete_params, "categoryid", lcl_categoryid)
    %>
      location.href="datamgr_categories_action.asp<%=lcl_delete_params%>";
  }
}

function validateFields(p_action) {
  var lcl_false_count              = 0;
  var lcl_return_merge_false_count = 0;
  var lcl_total_subcategories      = $('#totalsubcategories').val();
  var lcl_focus                    = '';
  var i                            = 0;

  if(lcl_total_subcategories > 0) {
     for(i = lcl_total_subcategories; i > 0; i--) {
         //if($('#sub_categoryname' + i).val() == '' && ! $('#sub_delete' + i).attr('checked')) {
         if($('#sub_categoryname' + i).val() == '' && ! $('#sub_delete' + i).prop('checked')) {
            $('#sub_categoryname' + i).val($('#sub_categoryname_original' + i).val());
            inlineMsg(document.getElementById('sub_categoryname' + i).id,'<strong>Required Field Missing: </strong>Sub-Category.',10,'sub_categoryname' + i);
            lcl_focus       = $('#sub_categoryname' + i);
            lcl_false_count = lcl_false_count + 1;
         } else {

            if(document.getElementById('mergeIntoCategory' + i)) {
               if($('#mergeIntoCategory' + i).val() != '') {
                  lcl_return_merge_false_count = mergeCategory(i);
                  lcl_false_count              = lcl_false_count + lcl_return_merge_false_count;
               }
            }

            clearMsg('sub_categoryname' + i);
         }
     }
  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("user_action").value = p_action;
     document.getElementById("categories_maint").submit();
     return true;
  }
}

function openWin(p_url, p_width, p_height) {
  w = 600;
  h = 400;

  if((p_width!="")&&(p_width!=undefined)) {
      w = p_width;
  }

  if((p_height!="")&&(p_height!=undefined)) {
      h = p_height;
  }

  l = (screen.availWidth/2)-(w/2);
  t = (screen.availHeight/2)-(h/2);

  lcl_url = p_url;

  eval('window.open("' + lcl_url + '", "_blank", "width=' + w + ',height=' + h + ',left=' + l + ',top=' + t + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,resizable=1")');
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
</script>

<style>
  .body {
     background-color: #ffffff;
  }

  #subcategory_buttons {
     margin-bottom: 50px;
  }

  #subcategory_add {
     position: relative;
     float:    left;
  }

  #subcategory_search {
     position: relative;
     float:    right;
  }

  .searchText {
     text-align: center;
     font-size:  16pt;
     color:      #800000;
  }

  .searchNote {
    color: #800000;
  }
</style>

</head>
<body class="body" onload="<%=lcl_onload%>">
<%
  response.write "  <form name=""categories_maint"" id=""categories_maint"" method=""post"" action=""datamgr_categories_action.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""screen_mode"" id=""screen_mode"" value=""" & lcl_screen_mode & """ size=""4"" maxlength=""4"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""user_action"" id=""user_action"" value="""" size=""4"" maxlength=""20"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""categoryid"" id=""categoryid"" value=""" & lcl_categoryid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""original_categoryid"" id=""original_categoryid"" value=""" & lcl_categoryid & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""dm_typeid"" id=""dm_typeid"" value=""" & lcl_dm_typeid & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ size=""4"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""sc_categoryname"" id=""sc_categoryname"" value=""" & lcl_sc_categoryname & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""isApproved"" id=""isApproved"" value=""" & lcl_isApproved & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""parent_categoryid"" id=""parent_categoryid"" value=""" & lcl_parent_categoryid & """ size=""5"" maxlength=""10"" />" & vbcrlf

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_pagetitle & ": " & lcl_screen_mode & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""Back to List"" class=""button"" onclick=""location.href='datamgr_categories_list.asp" & lcl_url_parameters & "';"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <p>" & vbcrlf
                            displayButtons "TOP", lcl_screen_mode, lcl_canDelete, lcl_return_parameters
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th colspan=""2"">&nbsp;</th>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap"">DM Type:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <span style=""color:#800000;"">" & lcl_displayDMT_description & "</span>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap"">Category:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <input type=""text"" name=""categoryname"" id=""categoryname"" value=""" & lcl_categoryname & """ size=""30"" maxlength=""100"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap"">Mappoint Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""mappointcolor"" id=""mappointcolor"">" & vbcrlf
                                        displayMapPointColors lcl_mappointcolor
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap"">&nbsp;</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""isActive"" id=""isActive"" value=""Y""" & lcl_checked_isactive & " /> Active" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap"">Created By:</td>" & vbcrlf
  response.write "                <td style=""color:#800000"">" & lcl_displayCreatedByInfo & "</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td nowrap=""nowrap"">Last Modified By:</td>" & vbcrlf
  response.write "                <td style=""color:#800000"">" & lcl_displayLastModifiedByInfo & "</td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "          </p>" & vbcrlf

 'BEGIN: Sub-Categories -------------------------------------------------------
  response.write "          <p>" & vbcrlf
  response.write "             <fieldset class=""fieldset"">" & vbcrlf
  response.write "               <legend>Sub-Categories&nbsp;</legend>" & vbcrlf
  response.write "               <div id=""subcategory_buttons"">" & vbcrlf
  response.write "                 <div id=""subcategory_add"">" & vbcrlf
  response.write "                   <input type=""button"" name=""sub_addSubCategoryButton"" id=""sub_addSubCategoryButton"" value=""Add Sub-Category"" class=""button"" />" & vbcrlf
  response.write "                 </div>" & vbcrlf

  if lcl_screen_mode = "EDIT" then
     response.write "                 <div id=""subcategory_search"">" & vbcrlf
     response.write "                   <input type=""text"" name=""sub_sc_categoryname"" id=""sub_sc_categoryname"" size=""30"" maxlength=""100"" />" & vbcrlf
     response.write "                   <input type=""button"" name=""sub_searchButton"" id=""sub_searchButton"" value=""Search Sub-Categories"" class=""button"" /><br />" & vbcrlf
     response.write "                 </div>" & vbcrlf
  end if

  lcl_sub_sc_categoryname = ""

  response.write "               </div>" & vbcrlf
  response.write "               <p>" & vbcrlf
                                    displaySubCategories session("orgid"), lcl_dm_typeid, lcl_categoryid, lcl_sub_sc_categoryname
  response.write "               </p>" & vbcrlf
  response.write "             </fieldset>" & vbcrlf
  response.write "          </p>" & vbcrlf
 'END: Sub-Categories ---------------------------------------------------------

                            displayButtons "BOTTOM", lcl_screen_mode, lcl_canDelete, lcl_return_parameters
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</form>" & vbcrlf

 'Determine if there are any scripts to run
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
  end if
%>

<!--#include file="../admin_footer.asp"-->

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displaySubCategories(iOrgID, iDMTypeID, iParentCategoryID, iSCCategoryName)

  response.write "<span id=""subcategories_results"">" & vbcrlf
  response.write "<table id=""subcategories_table"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""bottom"">" & vbcrlf
  response.write "      <th align=""left"">Sub-Category</th>" & vbcrlf
  response.write "      <th>Delete</th>" & vbcrlf
  'response.write "      <th>Created By</th>" & vbcrlf
  'response.write "      <th>Last Modified By</th>" & vbcrlf
  response.write "      <th>Approval Status</th>" & vbcrlf
  response.write "      <th>Approved/Denied By</th>" & vbcrlf
  response.write "      <th nowrap=""nowrap"">Merge Sub-Category into...</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  lcl_pcid = 0

  if iParentCategoryID <> "" then
     lcl_pcid = iParentCategoryID
     lcl_pcid = clng(lcl_pcid)
  end if

  if lcl_pcid > 0 then

     sSCCategoryName = ""

     if iSCCategoryName <> "" then
        sSCCategoryName = ucase(iSCCategoryName)
        sSCCategoryName = dbsafe(sSCCategoryName)
     end if

     sSQL = "SELECT dmc.categoryid, "
     sSQL = sSQL & " dmc.categoryname, "
     sSQL = sSQL & " dmc.isActive, "
     sSQL = sSQL & " dmc.isApproved, "
     'sSQL = sSQL & " dmc.createdbyid, "
     'sSQL = sSQL & " dmc.createdbydate, "
     'sSQL = sSQL & " dmc.lastmodifiedbyid, "
     'sSQL = sSQL & " dmc.lastmodifiedbydate, "
     sSQL = sSQL & " dmc.approvedeniedbyid, "
     sSQL = sSQL & " isnull(dmc.approvedeniedbydate,'') as approvedeniedbydate, "
     sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
     sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname, "
     sSQL = sSQL & " u3.firstname + ' ' + u3.lastname AS approvedeniedbyname "
     sSQL = sSQL & " FROM egov_dm_categories dmc "
     sSQL = sSQL &      " LEFT OUTER JOIN users u ON dmc.createdbyid = u.userid AND u.orgid = " & iOrgID
     sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON dmc.lastmodifiedbyid = u2.userid AND u2.orgid = " & iOrgID
     sSQL = sSQL &      " LEFT OUTER JOIN users u3 ON dmc.approvedeniedbyid = u3.userid AND u3.orgid = " & iOrgID
     sSQL = sSQL & " WHERE dmc.parent_categoryid = " & lcl_pcid

     if sSCCategoryName <> "" then
        sSQL = sSQL & " AND upper(dmc.categoryname) like ('%" & sSCCategoryName & "%') "
     end if

     set oGetSubCategories = Server.CreateObject("ADODB.Recordset")
     oGetSubCategories.Open sSQL, Application("DSN"), 3, 1

     if not oGetSubCategories.eof then
        lcl_bgcolor = "#ffffff"

        do while not oGetSubCategories.eof
           lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
     	   		iRowCount   = iRowCount + 1

          'Set up the display fields
           lcl_checked_sub_isApproved = ""

           if oGetSubCategories("isApproved") then
              lcl_checked_sub_isApproved = " checked=""checked"""
           end if

           'lcl_createdby_info     = formatAdminActionsInfo(oGetSubCategories("createdbyname"),       oGetSubCategories("createdbydate"))
           'lcl_lastmodified_info  = formatAdminActionsInfo(oGetSubCategories("lastmodifiedbyname"),  oGetSubCategories("lastmodifiedbydate"))

          'Set up Approve/Deny Buttons for display
           lcl_show_approvedButton        = 1
           lcl_show_deniedButton          = 1
           lcl_display_approvedDeniedInfo = ""
           lcl_approved_denied_status     = ""

           if not oGetSubCategories("isApproved") AND oGetSubCategories("approvedeniedbydate") = "1/1/1900" then
              lcl_approved_denied_status = "WAITING FOR<br />APPROVAL"
              lcl_approvedenied_info     = "&nbsp;"
           else
              lcl_approvedenied_info = formatAdminActionsInfo(oGetSubCategories("approvedeniedbyname"), oGetSubCategories("approvedeniedbydate"))

              if lcl_approvedenied_info <> "" then
                 if oGetSubCategories("isApproved") then
                    lcl_show_approvedButton    = 0
                    lcl_approved_denied_status = "APPROVED"
                 else
                    lcl_show_deniedButton      = 0
                    lcl_approved_denied_status = "DENIED"
                 end if

                 'lcl_display_approvedDeniedInfo = "<span class=""redText"">" & lcl_approved_denied_status & "</span><br />"
                 'lcl_display_approvedDeniedInfo = lcl_display_approvedDeniedInfo & lcl_approvedenied_info
              end if
           end if

           'response.write "  <tr id=""subcategoryrow" & oGetSubCategories("categoryid") & """ class=""subCategoryRow"" align=""center"" bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
           response.write "  <tr id=""subcategoryrow" & iRowCount & """ class=""subCategoryRow"" align=""center"" bgcolor=""" & lcl_bgcolor & """ valign=""top"">" & vbcrlf
           response.write "      <td align=""left"">" & vbcrlf
           response.write "          <input type=""hidden"" name=""sub_categoryid" & iRowCount & """ id=""sub_categoryid" & iRowCount & """ value=""" & oGetSubCategories("categoryid") & """ size=""3"" maxlength=""100"" />" & vbcrlf
           response.write "          <input type=""hidden"" name=""sub_categoryname_original" & iRowCount & """ id=""sub_categoryname_original" & iRowCount & """ value=""" & oGetSubCategories("categoryname") & """ size=""30"" maxlength=""100"" />" & vbcrlf
           response.write "          <input type=""text"" name=""sub_categoryname" & iRowCount & """ id=""sub_categoryname" & iRowCount & """ value=""" & oGetSubCategories("categoryname") & """ size=""30"" maxlength=""100"" onchange=""clearMsg('sub_categoryname" & iRowCount & "');"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "      <td>" & vbcrlf
           response.write "          <input type=""checkbox"" name=""sub_delete" & iRowCount & """ id=""sub_delete" & iRowCount & """ value=""Y"" onclick=""deleteCategory('" & iRowCount & "');"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
           'response.write "      <td nowrap=""nowrap"">" & lcl_createdby_info    & "</td>" & vbcrlf
           'response.write "      <td nowrap=""nowrap"">" & lcl_lastmodified_info & "</td>" & vbcrlf
'           response.write "      <td nowrap=""nowrap"">" & vbcrlf
'           response.write "          <span id=""approvedDeniedButtons" & iRowCount & """>" & vbcrlf

'           if lcl_show_approvedButton > 0 then
'              response.write "          <input type=""button"" name=""sub_approveButton" & iRowCount & """ id=""sub_approveButton" & iRowCount & """ class=""button"" value=""Approve"" onclick=""approveDenyCategory('" & iRowCount & "','A');"" />" & vbcrlf
'           end if

'           if lcl_show_deniedButton > 0 then
'              response.write "          <input type=""button"" name=""sub_denyButton" & iRowCount & """ id=""sub_denyButton" & iRowCount & """ class=""button"" value=""Deny"" onclick=""approveDenyCategory('" & iRowCount & "','D');"" />" & vbcrlf
'           end if

'           response.write "          </span>" & vbcrlf
'           response.write "      </td>" & vbcrlf
'           response.write "      <td nowrap=""nowrap""><span id=""approvedDeniedInfo" & iRowCount & """>" & lcl_display_approvedDeniedInfo & "</span></td>" & vbcrlf

        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <span id=""approvedDeniedStatus" & iRowCount & """ class=""redText"">" & lcl_approved_denied_status & "</span>" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td align=""center"" nowrap=""nowrap"">" & vbcrlf
        response.write "          <span id=""approvedDeniedInfo" & iRowCount & """>" & lcl_approvedenied_info & "</span><br />" & vbcrlf
        response.write "          <span id=""approvedDeniedButtons" & iRowCount & """>" & vbcrlf

        if lcl_show_approvedButton > 0 then
           response.write "          <input type=""button"" name=""approveButton" & iRowCount & """ id=""approveButton" & iRowCount & """ class=""button"" value=""Approve"" onclick=""approveDenyCategory('" & iRowCount & "','A');"" />" & vbcrlf
        end if

        if lcl_show_deniedButton > 0 then
           response.write "          <input type=""button"" name=""denyButton" & iRowCount & """ id=""denyButton" & iRowCount & """ class=""button"" value=""Deny"" onclick=""approveDenyCategory('" & iRowCount & "','D');"" />" & vbcrlf
        end if

        response.write "          </span>" & vbcrlf
        response.write "      </td>" & vbcrlf


           response.write "      <td>" & vbcrlf
           response.write "          <select name=""mergeIntoCategory" & iRowCount & """ id=""mergeIntoCategory" & iRowCount & """ onchange=""mergeCategory('" & iRowCount & "');"">" & vbcrlf
           response.write "            <option value=""""></option>" & vbcrlf
                                        lcl_parent_categoryid   = 0
                                        lcl_selected_categoryid = 0

                                        displayAllCategoriesOptions iOrgID, iDMTypeID, lcl_parent_categoryid, lcl_selected_categoryid
           response.write "          </select>" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "  </tr>"  & vbcrlf

           oGetSubCategories.movenext
        loop
     else
      		response.write "<p class=""norecords"">No Sub-Categories Available.</p>" & vbcrlf
     end if

     oGetSubCategories.close
     set oGetSubCategories = nothing
  end if

  response.write "</table>" & vbcrlf
  response.write "</span>" & vbcrlf
  response.write "<input type=""hidden"" name=""totalsubcategories"" id=""totalsubcategories"" value=""" & iRowCount & """ size=""5"" maxlength=""10"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iScreenMode, iCanDelete, iReturnParameters)

  if iTopBottom <> "" then
     iTopBottom = UCASE(iTopBottom)
  else
     iTopBottom = "TOP"
  end if

  if iTopBottom = "BOTTOM" then
     lcl_style_div = "margin-top:5px;"
  else
     lcl_style_div = "margin-bottom:5px;"
  end if

  response.write "<div style=""" & lcl_style_div & """>" & vbcrlf
  'response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='datamgr_categories_list.asp" & iReturnParameters & "'"" />" & vbcrlf

  if lcl_screen_mode = "ADD" then
     response.write "<input type=""button"" name=""addButton"" id=""addButton"" value=""Add"" class=""button"" onclick=""validateFields('ADD');"" />" & vbcrlf
  else
     if iCanDelete then
        response.write "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirmDelete();"" />" & vbcrlf
     end if

     response.write "<input type=""button"" name=""saveChangesButton"" id=""saveChangesButton"" value=""Save Changes"" class=""button"" onclick=""validateFields('UPDATE');"" />" & vbcrlf
  end if

  response.write "<div>" & vbcrlf

end sub

'-----------------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_value = REPLACE(p_value,"'","''")
  else
     lcl_value = p_value
  end if

  dbsafe = lcl_value

end function
%>