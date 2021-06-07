<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_categories_merge.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a DM Category
'
' MODIFICATION HISTORY
' 1.0 05/02/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

'Retrieve the parent categoryid in order to retrieve all of the sub-categories.
'If no value exists then redirect them back to the main results screen
 lcl_categoryid = ""

 if request("categoryid") <> "" then
    lcl_categoryid = request("categoryid")

    if not isnumeric(lcl_categoryid) then
response.write "here1"
'       response.redirect sLevel & "permissiondenied.asp"
    end if
 else
response.write "here2"
'    response.redirect sLevel & "permissiondenied.asp"
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
response.write "here3"
'      	response.redirect sLevel & "permissiondenied.asp"
    end if
 end if

 if not userhaspermission(session("userid"),lcl_feature) then
response.write "here4"
'   	response.redirect sLevel & "permissiondenied.asp"
 end if

 lcl_dm_typeid = clng(lcl_dm_typeid)
 lcl_pagetitle = getFeatureName(lcl_feature)
 lcl_pagetitle = lcl_pagetitle & " [Maintain Categories: Merge]"
 lcl_success   = request("success")

'Retrieve the search options
 lcl_sc_categoryname = ""

'Build return parameters
 lcl_url_parameters = ""
 lcl_url_parameters = setupUrlParameters(lcl_url_parameters, "dm_typeid",    lcl_dm_typeid)

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
 lcl_orgid              = session("orgid")
 lcl_categoryname       = ""
 lcl_isActive           = 1
 lcl_createdbyid        = 0
 lcl_createdbydate      = ""
 lcl_createdbyname      = ""
 lcl_lastmodifiedbyid   = 0
 lcl_lastmodifiedbydate = ""
 lcl_lastmodifiedbyname = ""
 lcl_parent_categoryid  = 0
 lcl_isApproved         = 1
 lcl_approvedbyid       = 0
 lcl_approvedbydate     = ""
 lcl_mappointcolor      = ""
 lcl_checked_isactive   = " checked=""checked"""

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
    sSQL = sSQL & " dmc.approvedbyid, "
    sSQL = sSQL & " dmc.approvedbydate, "
    sSQL = sSQL & " dmc.mappointcolor "
    sSQL = sSQL & " FROM egov_dm_categories dmc "
    sSQL = sSQL &      " LEFT OUTER JOIN users u ON dmc.createdbyid = u.userid AND u.orgid = " & session("orgid")
    sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON dmc.lastmodifiedbyid = u2.userid AND u2.orgid = " & session("orgid")
'    sSQL = sSQL & " WHERE dmc.categoryid = " & lcl_categoryid

    set oDMCategory = Server.CreateObject("ADODB.Recordset")
    oDMCategory.Open sSQL, Application("DSN"), 3, 1

    if not oDMCategory.eof then
       lcl_categoryid         = oDMCategory("categoryid")
       lcl_categoryname       = oDMCategory("categoryname")
       lcl_orgid              = oDMCategory("orgid")
       lcl_dm_typeid          = oDMCategory("dm_typeid")
       lcl_isActive           = oDMCategory("isActive")
       lcl_createdbyid        = oDMCategory("createdbyid")
       lcl_createdbydate      = oDMCategory("createdbydate")
       lcl_createdbyname      = oDMCategory("createdbyname")
       lcl_lastmodifiedbyid   = oDMCategory("lastmodifiedbyid")
       lcl_lastmodifiedbydate = oDMCategory("lastmodifiedbydate")
       lcl_lastmodifiedbyname = oDMCategory("lastmodifiedbyname")
       lcl_parent_categoryid  = oDMCategory("parent_categoryid")
       lcl_isApproved         = oDMCategory("isApproved")
       lcl_approvedbyid       = oDMCategory("approvedbyid")
       lcl_approvedbydate     = oDMCategory("approvedbydate")
       lcl_mappointcolor      = oDMCategory("mappointcolor")

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

'Get the description for the DM TypeID 
 lcl_displayDMT_description = getDMTypeDescription(lcl_dm_typeid)

'Check to see if this category has been associated to a DM Type.
'Know this will let us know if we can simply delete the "going to be merged" category
'or if we need to modify any DM Type category associations before deleting the "merged" category.
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
  <title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="layout_styles.css" />

 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/removespaces.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

  <script type="text/javascript" src="https://code.jquery.com/jquery-1.5.2.min.js"></script>
  <script type="text/javascript" src="https://github.com/jquery/jquery-ui.git"></script>

<script language="javascript">
$(document).ready(function() {

  $('#sub_searchButton').click(function() {
    var lcl_searchvalue = $('#sub_sc_categoryname').val();
    var lcl_foundCount  = 0;

    //Hide all of the rows
    $('.subCategoryRow').each(function() {

      //Get the "id" for the current <TR> in the loop
      var lcl_rowid = $(this).attr("id");

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
});

function saveDMChanges() {
  clearScreenMsg();
  $('#user_action').val('UPDATE');
  $('#datamgr_types_maint').attr('action','datamgr_action.asp');
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
  var lcl_false_count = 0;

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

  .fieldset {
     border: 1pt solid #808080;
     -webkit-border-radius: 5px;
     -moz-border-radius:    5px;
  }

  #categoryToMerge_fieldset {
    width:    50%;
    position: relative;
    float:    left;
  }

  #categoryMergeInto_fieldset {
    width:    50%;
    position: relative;
    float:    right;
  }

  #subcategory_buttons {
     margin-bottom: 50px;
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

  #screenMsg {
     color:       #ff0000;
     font-size:   10pt;
     font-weight: bold;
  }
</style>

</head>
<body class="body" onload="<%=lcl_onload%>">
<%
  response.write "  <form name=""categories_maint"" id=""categories_maint"" method=""post"" action=""datamgr_categories_action.asp"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""user_action"" id=""user_action"" value="""" size=""4"" maxlength=""20"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ />" & vbcrlf
  response.write "    <input type=""hidden"" name=""categoryid"" id=""categoryid"" value=""" & lcl_categoryid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""original_categoryid"" id=""original_categoryid"" value=""" & lcl_categoryid & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""dm_typeid"" id=""dm_typeid"" value=""" & lcl_dm_typeid & """ size=""5"" maxlength=""10"" />" & vbcrlf
  response.write "    <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ size=""4"" maxlength=""10"" />" & vbcrlf

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""10"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_pagetitle & "</strong></font><br />" & vbcrlf
  response.write "          <input type=""button"" name=""closeButton"" id=""closeButton"" value=""Close Window"" class=""button"" onclick=""parent.close();"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
  response.write "          <p>" & vbcrlf
                            displayButtons "TOP", lcl_canDelete, lcl_return_parameters

 'BEGIN: Sub-Categories -------------------------------------------------------
  response.write "          <p>" & vbcrlf
  response.write "             <div id=""toMerge_fieldset"">" & vbcrlf
  response.write "             <fieldset class=""fieldset"">" & vbcrlf
  response.write "               <legend>Sub-Category to be Merged&nbsp;</legend>" & vbcrlf
  response.write "               <p>" & vbcrlf
  response.write "                  Select a Parent Category to :"
  response.write "                  <select name=""defaultCategoryID"" id=""defaultCategoryID"" onchange=""clearMsg('defaultCategoryID');"">" & vbcrlf
  response.write "                    <option value=""0""></option>" & vbcrlf
                                      lcl_parent_categoryid = 0

                                      displayDMTCategories session("orgid"), lcl_dm_typeid, lcl_parent_categoryid, lcl_defaultcategoryid
  response.write "                  </select>" & vbcrlf
  response.write "               </p>" & vbcrlf

  response.write "               <p>" & vbcrlf
                                    lcl_sub_sc_categoryname = ""

'                                    displaySubCategories session("orgid"), lcl_categoryid, lcl_sub_sc_categoryname
  response.write "               </p>" & vbcrlf
  response.write "             </fieldset>" & vbcrlf
  response.write "             </div>" & vbcrlf
  response.write "             <div id=""categoryMergeInto_fieldset"">" & vbcrlf
  response.write "             <fieldset class=""fieldset"">" & vbcrlf
  response.write "               <legend>Sub-Category to be Merged&nbsp;</legend>" & vbcrlf
  response.write "               <p>" & vbcrlf
  response.write "                  Select a Parent Category to :"
  response.write "                  <select name=""defaultCategoryID"" id=""defaultCategoryID"" onchange=""clearMsg('defaultCategoryID');"">" & vbcrlf
  response.write "                    <option value=""0""></option>" & vbcrlf
                                      lcl_parent_categoryid = 0

                                      displayDMTCategories session("orgid"), lcl_dm_typeid, lcl_parent_categoryid, lcl_defaultcategoryid
  response.write "                  </select>" & vbcrlf
  response.write "               </p>" & vbcrlf

  response.write "               <p>" & vbcrlf
                                    lcl_sub_sc_categoryname = ""

'                                    displaySubCategories session("orgid"), lcl_categoryid, lcl_sub_sc_categoryname
  response.write "               </p>" & vbcrlf
  response.write "             </fieldset>" & vbcrlf
  response.write "             </div>" & vbcrlf
  response.write "          </p>" & vbcrlf
 'END: Sub-Categories ---------------------------------------------------------

                            displayButtons "BOTTOM", lcl_canDelete, lcl_return_parameters
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
sub displaySubCategories(iOrgID, iParentCategoryID, iSCCategoryName)

  response.write "<span id=""subcategories_results"">" & vbcrlf
  response.write "<table id=""subcategories_table"" cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr valign=""bottom"">" & vbcrlf
  response.write "      <th align=""left"">Sub-Category</th>" & vbcrlf
  response.write "      <th>Delete</th>" & vbcrlf
  response.write "      <th>Created By</th>" & vbcrlf
  response.write "      <th>Last Modified By</th>" & vbcrlf
  response.write "      <th>&nbsp;</th>" & vbcrlf
  response.write "      <th>Approved/Denied By</th>" & vbcrlf
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
     sSQL = sSQL & " dmc.createdbyid, "
     sSQL = sSQL & " dmc.createdbydate, "
     sSQL = sSQL & " dmc.lastmodifiedbyid, "
     sSQL = sSQL & " dmc.lastmodifiedbydate, "
     sSQL = sSQL & " dmc.approvedbyid, "
     sSQL = sSQL & " dmc.approvedbydate, "
     sSQL = sSQL & " u.firstname + ' ' + u.lastname AS createdbyname, "
     sSQL = sSQL & " u2.firstname + ' ' + u2.lastname AS lastmodifiedbyname, "
     sSQL = sSQL & " u3.firstname + ' ' + u3.lastname AS approvedbyname "
     sSQL = sSQL & " FROM egov_dm_categories dmc "
     sSQL = sSQL &      " LEFT OUTER JOIN users u ON dmc.createdbyid = u.userid AND u.orgid = " & iOrgID
     sSQL = sSQL &      " LEFT OUTER JOIN users u2 ON dmc.lastmodifiedbyid = u2.userid AND u2.orgid = " & iOrgID
     sSQL = sSQL &      " LEFT OUTER JOIN users u3 ON dmc.approvedbyid = u3.userid AND u3.orgid = " & iOrgID
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

           lcl_createdby_info    = formatAdminActionsInfo(oGetSubCategories("createdbyname"), oGetSubCategories("createdbydate"))
           lcl_lastmodified_info = formatAdminActionsInfo(oGetSubCategories("lastmodifiedbyname"), oGetSubCategories("lastmodifiedbydate"))
           lcl_approval_info     = formatAdminActionsInfo(oGetSubCategories("approvedbyname"), oGetSubCategories("approvedbydate"))
           'lcl_denied_info       = formatAdminActionsInfo(oGetSubCategories("deniedbyname"), oGetSubCategories("deniedbydate"))
           lcl_denied_info = "denied info"

          'Set up Approve/Deny Buttons for display
           if lcl_approval_info <> "" OR lcl_denied_info <> "" then
              lcl_show_approvedButton        = 0
              lcl_show_deniedButton          = 0
              lcl_display_approvedDeniedInfo = ""

              if oGetSubCategories("isApproved") then
                 lcl_show_deniedButton = lcl_show_deniedButton + 1

                 if lcl_approval_info <> "" then
                    lcl_display_approvedDeniedInfo = lcl_approval_info
                 else
                    lcl_show_approvedButton = lcl_show_approvedButton + 1
                 end if
              else
                 lcl_show_approvedButton = lcl_show_approvedButton + 1

                 if lcl_denied_info <> "" then
                    lcl_display_approvedDeniedInfo = lcl_denied_info
                 else
                    lcl_show_deniedButton = lcl_show_deniedButton + 1
                 end if
              end if
           end if

           'response.write "  <tr id=""subcategoryrow" & oGetSubCategories("categoryid") & """ class=""subCategoryRow"" align=""center"" bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
           response.write "  <tr id=""subcategoryrow" & iRowCount & """ class=""subCategoryRow"" align=""center"" bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" valign=""top"">" & vbcrlf
           response.write "      <td align=""left"">" & vbcrlf
           response.write "          <input type=""hidden"" name=""sub_categoryid" & iRowCount & """ id=""sub_categoryid" & iRowCount & """ value=""" & oGetSubCategories("categoryid") & """ size=""3"" maxlength=""100"" />" & vbcrlf
           response.write "          <input type=""text"" name=""sub_categoryname" & iRowCount & """ id=""sub_categoryname" & iRowCount & """ value=""" & oGetSubCategories("categoryname") & """ size=""30"" maxlength=""100"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "      <td>" & vbcrlf
           response.write "          <input type=""checkbox"" name=""sub_delete" & iRowCount & """ id=""sub_delete" & iRowCount & """ value=""Y"" />" & vbcrlf
           response.write "      </td>" & vbcrlf
           response.write "      <td nowrap=""nowrap"">" & lcl_createdby_info    & "</td>" & vbcrlf
           response.write "      <td nowrap=""nowrap"">" & lcl_lastmodified_info & "</td>" & vbcrlf
           response.write "      <td nowrap=""nowrap"">" & vbcrlf

           if lcl_show_approvedButton > 0 then
              response.write "          <input type=""button"" name=""sub_approveButton" & iRowCount & """ id=""sub_approveButton" & iRowCount & """ class=""button"" value=""Approve"" onclick=""alert('approved');"" />" & vbcrlf
           end if

           if lcl_show_deniedButton > 0 then
              response.write "          <input type=""button"" name=""sub_denyButton" & iRowCount & """ id=""sub_denyButton" & iRowCount & """ class=""button"" value=""Deny"" onclick=""alert('denied');"" />" & vbcrlf
           end if

           response.write "      </td>" & vbcrlf
           response.write "      <td nowrap=""nowrap"">" & lcl_display_approvedDeniedInfo & "</td>" & vbcrlf
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
  response.write "<input type=""text"" name=""totalsubcategories"" id=""totalsubcategories"" value=""" & iRowCount & """ size=""5"" maxlength=""10"" />" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom, iCanDelete, iReturnParameters)

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
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='datamgr_list.asp" & iReturnParameters & "'"" />" & vbcrlf

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
