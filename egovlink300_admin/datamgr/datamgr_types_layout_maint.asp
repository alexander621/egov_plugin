<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="datamgr_global_functions.asp" //-->
<!-- #include file="datamgr_build_sections_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: datamgr_types_layout_maint.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin to modify a DM Type Layout
'
' MODIFICATION HISTORY
' 1.0 02/09/2011 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("datamgr") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel               = "../"  'Override of value from common.asp
 lcl_onload           = ""
 lcl_isRootAdmin      = False

'Determine if there is a specific feature associated to the DM Types
'  "Is Limited": means that the admin user is maintaining a specific feature instead of the root admin viewing ALL MapPoint Types
'      (i.e. "Is Limited" will be true when an admin clicks on "Maintain Available Properties (fields)", but NOT 
'            when a root admin clicks on "Maintain MapPoint Types" and then selects a specific MapPoint Type to edit.
 lcl_dm_typeid    = ""
 lcl_feature      = "datamgr_types_maint"
 lcl_isLimited    = False
 lcl_pagetitle    = "DM Types"
 lcl_sectiontitle = ": Maintain Layout"

 if request("f") <> "" AND request("f") <> "datamgr_types_maint" then
    lcl_feature   = request("f")
    lcl_isLimited = True
    lcl_pagetitle = getFeatureName(lcl_feature)
    lcl_pagetitle = replace(lcl_pagetitle," (Fields)","")
    lcl_pagetitle = replace(lcl_pagetitle,"Maintain ","")

   'Retrieve the dm_typeid
    if request("dm_typeid") <> "" then
       lcl_dm_typeid = request("dm_typeid")
    else
       lcl_dm_typeid = getDMTypeByFeature(session("orgid"), "feature_maintain_fields", lcl_feature)

       if lcl_dm_typeid = 0 then
         	response.redirect sLevel & "permissiondenied.asp"
       end if
    end if
 else
    if request("dm_typeid") <> "" then
       lcl_dm_typeid = request("dm_typeid")
    end if
 end if

'Build the section title
 lcl_sectiontitle = lcl_pagetitle & lcl_sectiontitle

 if not userhaspermission(session("userid"),lcl_feature) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

'Determine if the user is a "root admin"
 if UserIsRootAdmin(session("userid")) then
    lcl_isRootAdmin = True
 end if

'Retrieve the dm_typeid to be maintain.
'If no value exists AND the screen_mode does not equal ADD then redirect them back to the main results screen
 'if request("dm_typeid") <> "" then

 if lcl_dm_typeid <> "" then
    if not isnumeric(lcl_dm_typeid) then
       response.redirect "datamgr_types_maint.asp?dm_typeid=" & lcl_dm_typeid
    end if
 else
    lcl_dm_typeid = 0
 end if

'Set up local variables
 lcl_layoutid = 0

 if request("layoutid") <> "" then
    lcl_layoutid = request("layoutid")
 end if

'Determine if the Layout exists and if not then get the "Original" Layout
 getLayoutInfo lcl_layoutid, lcl_layoutname, lcl_isOriginalLayout, lcl_useLayoutSections, lcl_totalcolumns, _
               lcl_columnwidth_left, lcl_columnwidth_middle, lcl_columnwidth_right

'Check for a screen message
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = lcl_onload & "displayScreenMsg('" & lcl_msg & "');"

    if lcl_success = "SU" then
       lcl_onload = lcl_onload & "window.opener.location.reload();"
    end if
 end if

 dim lcl_scripts
%>
<html>
<head>
  <title>E-Gov Administration Console {<%=lcl_sectiontitle%>}</title>

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="layout_styles.css" />
  <link href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8/themes/base/jquery-ui.css" rel="stylesheet" type="text/css"/>

 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/selectAll.js"></script>
  <script language="javascript" src="../scripts/textareamaxlength.js"></script>
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
  <script language="javascript" src="../scripts/datamgr_fields_addrow.js"></script>

<% '  <script type="text/javascript" src="../scripts/jquery-1.4.4.min.js"></script> %>

  <script type="text/javascript" src="../scripts/jquery-1.6.1.min.js"></script>
  <script type="text/javascript" src="../scripts/jquery-ui-1.8.4.custom.min.js"></script>

<style type="text/css">
  .hidden            { display: none; }
  .requiredFieldsMsg { color: #800000; }

  .instructions, ol, li {
     /*color:     #800000;*/
     font-size: 11px;
  }

  .body {
     background-color: #ffffff;
     margin-top:       10px;
  }
</style>

<script language="javascript">

$(document).ready(function(){
  $.fn.enableDisableButton = function(iCompareField1, iCompareField2, iButtonID, iEvent, iAction1, iAction2) { 
    if($(iCompareField1).val() != $(iCompareField2).val()) {
       //$(iButtonID).attr('disabled',false);
       //$(iButtonID).attr(iEvent,iAction1);
       $(iButtonID).prop(iEvent,iAction1);
    } else {
       //$(iButtonID).attr('disabled',true);
       //$(iButtonID).attr(iEvent,iAction2);
       $(iButtonID).prop(iEvent,iAction2);
    }
  }

  //Setup for drag-n-drop sections
  $('.column').sortable({
     connectWith:          '.column',
     handle:               'h2',
     cursor:               'move',
     placeholder:          'placeholder',
     forcePlaceholderSize: true,
     opacity:              0.4,
     stop: function(event, ui) {
        $(ui.item).find('h2').click();

        //Reorder the sections (dragboxes)
      		var sortorder        = '';
        var lcl_totalcolumns = 1;

        if(document.getElementById("totalcolumns").value != '') {
           lcl_totalcolumns = document.getElementById("totalcolumns").value;
        }

	      	$('.column').each(function(){
          var itemorder           = $(this).sortable('toArray');
          var lcl_total_items     = itemorder.length;
          var lcl_sectionid       = '';
          var lcl_dm_typeid       = '';
          var lcl_dm_sectionid    = '';
          var lcl_sectionlocation = 'L';
          var lcl_dragboxtype     = '';
          var lcl_id              = '';
          var lcl_sectionactive   = 'Y';
          var sParameter          = '';

          for(var i = 0; i < lcl_total_items; i++) {
              //Depending on the section area we are working with determines which value we try and "replace" so
              //we can find the actual section id
              //dmt_section = active sections assigned to a DM Type
              //section     = active, non-assigned to a DM Type sections
              lcl_dm_typeid = document.getElementById("dm_typeid").value;

              lcl_sectionid = itemorder[i];
              lcl_sectionid = lcl_sectionid.replace('dragbox','');

              //if(lcl_sectionid.indexOf('dmt_section_' + lcl_dm_typeid) > -1) {
              //   lcl_dragboxtype = 'dmt_section_' + lcl_dm_typeid + '_';
              //} else {
              //   lcl_dragboxtype = 'section_';
              //}

              //lcl_sectionid    = lcl_sectionid.replace(lcl_dragboxtype,"");
              //lcl_dm_sectionid = $('#' + lcl_dragboxtype + 'dm_sectionid_' + lcl_sectionid).val();
              lcl_dm_sectionid = $('#dm_sectionid_' + lcl_sectionid).val();

              //determine which column this dragbox is in
              //if(lcl_totalcolumns == 4) {
              //   if(this.id == "column4") {
              //      lcl_sectionactive = 'N';
              //   }else if(this.id == "column3") {
              //      lcl_sectionlocation = 'R';
              //   }else if(this.id == "column2") {
              //      lcl_sectionlocation = 'M';
              //   }
              //} else if(lcl_totalcolumns == 3) {
              //   if(this.id == "column2") {
              //      lcl_sectionlocation = 'M';
              //   } else if(this.id == "column3") {
              //      lcl_sectionlocation = 'R';
              //   }
              //} else if(lcl_totalcolumns == 2) {
              //   if(this.id == "column2") {
              //      lcl_sectionlocation = 'R';
              //   }
              //}

              //Determine which column THIS dragbox is in.
              //The "Available Sections" column number changes based on the number of columns
              //  in the layout.
              //NOTE: it is understood that there is at least ONE active column and ONE inactive column
              lcl_lastcolumn         = lcl_totalcolumns;
              lcl_totalActiveColumns = (lcl_totalcolumns - 1);

              if(lcl_totalActiveColumns == 3) {
                 if(this.id == "column2") {
                    lcl_sectionlocation = 'M';
                 } else if(this.id == "column3") {
                    lcl_sectionlocation = 'R';
                 }
              } else if(lcl_totalActiveColumns == 2) {
                 if(this.id == "column2") {
                    lcl_sectionlocation = 'R';
                 }
              }

              if(this.id == "column" + lcl_lastcolumn) {
                 lcl_sectionactive = 'N';
              }

              //update the section order and location input fields
              $('#sectionorder_'    + lcl_sectionid).val(i+1);
              $('#sectionlocation_' + lcl_sectionid).val(lcl_sectionlocation);
              $('#sectionactive_'   + lcl_sectionid).val(lcl_sectionactive);

              //sParameter  = 'orgid='            + encodeURIComponent('<%=session("orgid")%>');
              //sParameter += '&dm_typeid=' + encodeURIComponent(lcl_dm_typeid);
              //sParameter += '&dm_sectionid='    + encodeURIComponent(lcl_dm_sectionid);
              //sParameter += '&sectionid='       + encodeURIComponent(lcl_sectionid);
              //sParameter += '&sectionlocation=' + encodeURIComponent(lcl_sectionlocation);
              //sParameter += '&sectionorder='    + encodeURIComponent(i+1);
              //sParameter += '&sectionactive='   + encodeURIComponent(lcl_sectionactive);
              //sParameter += '&isAjax=Y';
    //alert('left off here.  need to change the ajax call to a form-call for saving changes. this issue with the ajax call is that it cannot handle NEW dm_sectionids to MPT.');
              //doAjax('update_dmt_section.asp', sParameter, 'displayScreenMsg', 'post', '0');
          }
        });
     }
  })
  .disableSelection();

  //enable/disable the "save layout" button when the page is first opened.
  $('#layoutid').enableDisableButton('#layoutid', '#original_layoutid', '#changeLayoutButton', 'disabled', false, true);
  $('#saveButton').enableDisableButton('#layoutid', '#original_layoutid', '#saveButton', 'disabled', true, false);

  //enable/disable "save layout" button
  $('#layoutid').change(function() {
    $('#layoutid').enableDisableButton('#layoutid', '#original_layoutid', '#changeLayoutButton', 'disabled', false, true);
    $('#saveButton').enableDisableButton('#layoutid', '#original_layoutid', '#saveButton', 'disabled', true, false);
  })
//  change();
});

function changeLayout() {
  lcl_dm_typeid = document.getElementById("dm_typeid")
  lcl_layoutid  = document.getElementById("layoutid");
  lcl_orgid     = document.getElementById("orgid")

  document.getElementById("user_action").value = "CHANGE_LAYOUT";
  document.getElementById("dmt_layout_maint").submit();
}

function saveChanges() {
  document.getElementById("user_action").value = "SAVE_CHANGES";
  document.getElementById("dmt_layout_maint").submit();
}

function confirmDelete() {
  //var r = confirm('Are you sure you want to delete the "' + document.getElementById("title").value + '" blog entry?  \r NOTE: Any/All comments will be deleted as well.');
  var r = confirm('Are you sure you want to delete: "' + document.getElementById("description").value + '"');
  if (r==true) {
      location.href="datamgr_types_action.asp?user_action=DELETE&dm_typeid=<%=lcl_dm_typeid & lcl_isTemplate_url%>";
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
</script>

</head>
<body class="body" onload="<%=lcl_onload%>">
<%
  response.write "<form name=""dmt_layout_maint"" id=""dmt_layout_maint"" method=""post"" action=""datamgr_types_layout_action.asp"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""dm_typeid"" id=""dm_typeid"" value=""" & lcl_dm_typeid & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""user_action"" id=""user_action"" value="" size=""4"" maxlength=""20"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orgid"" id=""orgid"" value=""" & session("orgid") & """ size=""4"" maxlength=""10"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""f"" id=""f"" value=""" & lcl_feature & """ size=""10"" maxlength=""50"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""t"" id=""t"" value=""" & request("t") & """ size=""5"" maxlength=""5"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""original_layoutid"" id=""original_layoutid"" value=""" & lcl_layoutid & """ size=""5"" maxlength=""10"" />" & vbcrlf

  response.write "<div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""800"" class=""start"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <font size=""+1""><strong>" & lcl_sectiontitle & "</strong></font>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td nowrap=""nowrap"">" & vbcrlf
  response.write "          Layout: " & vbcrlf
  response.write "          <select name=""layoutid"" id=""layoutid"">" & vbcrlf
                              displayLayoutOptions lcl_layoutid
  response.write "          </select>" & vbcrlf
  response.write "          <input type=""button"" name=""changeLayoutButton"" id=""changeLayoutButton"" value=""Change Layout"" class=""button"" onclick=""changeLayout()"" />" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td style=""padding-left:10px"" rowspan=""2"">" & vbcrlf
  response.write "          <div class=""instructions"">" & vbcrlf
  response.write "            <strong>Instructions: </strong>To reposition sections in a DM Type Layout click on section header "
  response.write "            and drag the section to the desired column. Drop the section when it is in the position you want "
  response.write "            to move it to. ALL ""section-field(s)"" associated to the section will:<br />" & vbcrlf
  response.write "            <ol>" & vbcrlf
  response.write "                <li>Be automatically added and enabled to the DM Type when a section is placed into the ""Current Layout"".</li>" & vbcrlf
  response.write "                <li>The section and its field(s) will be automatically disabled for the DM Type when a section "
  response.write "                    is removed from the ""Current Layout"" and dropped back into the ""Available Sections"" area.</li>" & vbcrlf
  response.write "            </ol>" & vbcrlf
  response.write "          </div>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""bottom"">" & vbcrlf
  response.write "      <td>" & vbcrlf
                            displayButtons "TOP"
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "  <tr valign=""top"">" & vbcrlf
  response.write "      <td colspan=""2"">" & vbcrlf
                           'Build the Layout
                           'Retrieve any/all fields related to this DM Type
                           'ONLY show these field if "screen mode" = "EDIT"
                            lcl_displayFieldsetLegend    = True
                            lcl_displayFieldsetBorders   = True
                            lcl_displayAvailableSections = True
                            lcl_screen_mode              = "DRAG"
                            lcl_dmid                     = ""

                            buildDMLayout lcl_layoutid, lcl_dm_typeid, lcl_dmid, lcl_displayFieldsetLegend, _
                                          lcl_displayFieldsetBorders, lcl_displayAvailableSections, lcl_screen_mode

                           'Display the bottom row of buttons
                            displayButtons "BOTTOM"

  response.write "          </p>" & vbcrlf
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
<%
response.write "</body>" & vbcrlf
response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayButtons(iTopBottom)

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
  response.write "  <input type=""button"" name=""closeButton"" id=""closeButton"" value=""Close Window"" class=""button"" onclick=""parent.close();"" />" & vbcrlf
  response.write "  <input type=""button"" name=""saveButton"" id=""saveButton"" value=""Save Changes"" class=""button"" onclick=""saveChanges()"" />" & vbcrlf
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