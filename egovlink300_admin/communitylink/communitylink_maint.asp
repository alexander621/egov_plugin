<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="communitylink_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: communitylink_maint.asp
' AUTHOR:    David Boyer
' CREATED:   04/14/09
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module allows an admin user to maintain the look-n-feel of their org's CommunityLink page.
'
' MODIFICATION HISTORY
' 1.0 04/14/09 David Boyer - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("communitylink") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not userhaspermission(session("userid"),"communitylink_maint") then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 set oOrg = New classOrganization

 dim lcl_scripts

 lcl_pagetitle = "Maintain CommunityLink"

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check to see if a custom website size has been entered and needs to be shown/hidden.
 lcl_onload = lcl_onload & "setupCustomSize();"

'Check for a CommunityLink record for the org.
'If one DOES exist then pull all of the values.
'If one does NOT exist then create it and enter the default values.
 lcl_communitylinkid = getCommunityLinkID(session("orgid"), _
                                          session("userid"))

'Retrieve the CommunityLink record.
 getCommunityLinkInfo lcl_communitylinkid, _
                      session("orgid"), _
                      lcl_isEgovHomePage, _
                      lcl_website_size, _
                      lcl_website_size_customsize, _
                      lcl_website_alignment, _
                      lcl_website_bgcolor, _
                      lcl_showlogo, _
                      lcl_logo_filename, _
                      lcl_logo_filenamebg, _
                      lcl_logo_alignment, _
                      lcl_showtopbar, _
                      lcl_topbar_bgcolor, _
                      lcl_topbar_fonttype, _
                      lcl_topbar_fontcolor, _
                      lcl_topbar_fontcolorhover, _
                      lcl_showsidemenubar, _
                      lcl_sidemenubar_alignment, _
                      lcl_sidemenuoption_bgcolor, _
                      lcl_sidemenuoption_bgcolorhover, _
                      lcl_sidemenuoption_alignment, _
                      lcl_sidemenuoption_fonttype, _
                      lcl_sidemenuoption_fontcolor, _
                      lcl_sidemenuoption_fontcolorhover, _
                      lcl_showpageheader, _
                      lcl_pageheader_alignment, _
                      lcl_pageheader_fontsize, _
                      lcl_pageheader_fontcolor, _
                      lcl_pageheader_fonttype, _
                      lcl_pageheader_bgcolor, _
                      lcl_showfooter, _
                      lcl_footer_bgcolor, _
                      lcl_footer_fonttype, _
                      lcl_footer_fontcolor, _
                      lcl_footer_fontcolorhover, _
                      lcl_showRSS, _
                      lcl_url_twitter, _
                      lcl_url_facebook, _
                      lcl_url_myspace, _
                      lcl_url_blogger

'Check for org features
 lcl_orghasfeature_administrationlink = orghasfeature("AdministrationLink")
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>
<!--  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" /> -->

	 <link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" href="../global.css" />
  <link rel="stylesheet" href="../custom/css/tooltip.css" />
<!-- 	<link rel="stylesheet" href="../yui/build/tabview/assets/skins/sam/tabview.css" /> -->
  <!-- <link rel="stylesheet" type="text/css" href="../custom/css/dragdrop.css" /> -->

  <script src="../scripts/modules.js"></script>
 	<script src="../scripts/ajaxLib.js"></script>
  <script src="../scripts/tooltip_new.js"></script>
  <script src="../scripts/formvalidation_msgdisplay.js"></script>
  <!-- <script language="javascript" src="../scripts/drag_drop.js"></script> -->
	
	<!--
 	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js"></script>
 	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>
 	<script type="text/javascript" src="../yui/build/tabview/tabview.js"></script>
	-->


  <script src="../scripts/jquery-1.9.1.min.js"></script>

<style>
  #content
  {
     width: 90%;
  }

  #content table
  {
     background: none;
  }  

  .fieldset
  {
     margin: 10px 0px;
     border-radius: 6px;
     background-color: #eeeeee;
  }

  .fieldset legend
  {
     background-color: #ffffff;
     padding: 4px 8px;
     border: 1pt solid #808080;
     border-radius: 6px;
     color: #800000;
     font-size: 1.25em;
  }

  #tableLayoutOptions
  {
     margin: 5px 0px;
  }

  #tableLayoutOptions td
  {
     background-color: #eeeeee;
  }

  .closeButton { 
     cursor: pointer; 
  } "

  .helpOption     { cursor: pointer } 

  .helpOptionText { 
     background-color:      #a80000; 
     font-size:             12px; 
     color:                 #ffffff; 
     padding:               5px 5px; 
     margin:                5px 5px; 
     border:                1pt solid #000000; 
     -webkit-border-radius: 5px; 
     -moz-border-radius:    5px; 
  } 

 /* BEGIN: Community Link Options ------------------------------------------ */
  .communitylink_table { 
     border: 1pt solid #000000; 
  } 

  .communitylink_table th { 
     font-weight:      bold; 
     border-bottom:    1pt solid #000000; 
     background-color: #cccccc; 
  } 

  .communityLink_viewLinks:link, 
  .communityLink_viewLinks:visited, 
  .communityLink_viewLinks:active { 
     color: #800000; 
  } 

  .communityLink_viewLinks:hover { 
     color: #800000; 
     text-decoration: underline; 
  } 

  .communityLink_bottomborder { 
     border-bottom: 1pt dotted #808080; "  & vbcrlf
  } 

  .sectionlinks_viewall_url
  {
     margin-bottom: 4px;
  }
 /* END: Community Link Options -------------------------------------------- */

</style>

<script>
<!--
//  var tabView;

//  (function() {
//  	 tabView = new YAHOO.widget.TabView('demo');
//  	 tabView.set('activeIndex', 0); 

//  })();

$(document).ready(function(){
  $('.sectionTable').css('display','none');
  //$('.fieldset').css('display','none');
  $('[id^="sectionFieldset_"]').css('display','none');
  $('.queryfilter_div').css('display','none');
  $('.closeButton').css('display','none');
  $('.helpOptionText').hide();
  $('.sectionlinks_viewall_url').attr('disabled','true');
});

function showHideOptions(iAction, iType, iRowCount) {
  var lcl_action;;

  if((iAction == '') || (iAction == undefined)) {
     lcl_action = 'H';
  } else {
     lcl_action = iAction;
  }

  if(lcl_action == 'S') {
     $('#edit'         + iType + '_' + iRowCount).val('Y');
     $('#editButton_'  + iType + '_' + iRowCount).prop('disabled',true);
     $('#closeButton_' + iType + '_' + iRowCount).show('show', function() {
       $('#queryfilter_div_' + iRowCount).show('slow',function() {
         if(iType == 'CL') {
            $('#sectionTable_' + iRowCount).show('slow');
         }

         $('#sectionFieldset_' + iType + '_' + iRowCount).show('slow', function() {
            enableDisableViewALLURL(iType + '_' + iRowCount, 'N');
         });
       });
     });
  } else {
     $('#edit'         + iType + '_' + iRowCount).val('N');
     $('#editButton_'  + iType + '_' + iRowCount).prop('disabled','');
     $('#closeButton_' + iType + '_' + iRowCount).hide('slow', function() {

       if(iType == 'CL') {
          $('#sectionTable_' + iRowCount).hide('slow');
          if($('#editSAVVY_' + iRowCount).val() == 'N') {
             $('#queryfilter_div_' + iRowCount).hide('slow');
          }
       } else {
          if($('#editCL_' + iRowCount).val() == 'N') {
             $('#queryfilter_div_' + iRowCount).hide('slow');
          }
       }

       $('#sectionFieldset_' + iType + '_' + iRowCount).hide('slow');
     });
  }
}

function showHideHelp(iType, iRowCount) {
  $('#helpFeature_edit_' + iType + '_' + iRowCount + '_text').toggle('slow');
}

function enableDisableOptions(iType,iRowCount) {
  lcl_showSection = document.getElementById("showSection_" + iType + "_" + iRowCount);

  if(lcl_showSection.checked) {
     lcl_disabled  = false;
     lcl_displayed = "inline";
  }else{
     lcl_disabled  = true;
     lcl_displayed = "none";
  }

  document.getElementById("styleProperties_" + iType + "_" + iRowCount).style.display = lcl_displayed;

  if(iType == "CL") {
     document.getElementById("styleProperties_Portal_"         + iRowCount).style.display = lcl_displayed;
     document.getElementById("styleProperties_DisplayOrder_"   + iRowCount).style.display = lcl_displayed;
     document.getElementById("styleProperties_RSSFeedFeature_" + iRowCount).style.display = lcl_displayed;
  }

  //If either/both of the CommunityLink AND Savvy/IFRAME "display" checkbox(es) are "checked" then:
  //1. enabled the Feature Name input field.
  //2. enable the RESET button for the row.
  if(document.getElementById("showSection_CL_" + iRowCount).checked || document.getElementById("showSection_SAVVY_" + iRowCount).checked) {
     document.getElementById("featurename_" + iRowCount).disabled = false;
     document.getElementById("styleProperties_ResetButton_"  + iRowCount).style.display = "block";
  }else{
     document.getElementById("featurename_" + iRowCount).disabled = true;
     document.getElementById("styleProperties_ResetButton_"  + iRowCount).style.display = "none";
  }
}

function confirm_delete(iFeedID, iTotalItems) {
  var lcl_rssTitle = document.getElementById("rssfeed"+iFeedID).innerHTML;

  if(iTotalItems > 0) {
     lcl_msg  = '"' + lcl_rssTitle + '" cannot be deleted as there are RSS Items associated to it.\n';
     lcl_msg += 'Set the RSS Feed to "inactive".';

     alert(lcl_msg);
  }else{
    	if (confirm("Are you sure you want to delete '" + lcl_rssTitle + "' ?")) { 
  	   			//DELETE HAS BEEN VERIFIED
   		  		location.href='rssfeeds_action.asp?user_action=DELETE&feedid='+ iFeedID;
     }
		}
}

function setupCustomSize() {
  lcl_size = document.getElementById("website_size").value;

  if(lcl_size == "C") {
     lcl_display = "inline";
  }else{
     lcl_display = "none";
  }

  document.getElementById("website_size_customsize_span").style.display = lcl_display;
}
<%
 'Setup the valid image types
  lcl_imgTypesDisplay = "BMP,GIF,JPG,JPEG,PNG,TIF"

  lcl_imgTypes = ""
  lcl_imgTypes = lcl_imgTypes & "(lcl_ext==""BMP"")"
  lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""GIF"")"
  lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""JPG"")"
  lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""JPEG"")"
  lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""PNG"")"
  lcl_imgTypes = lcl_imgTypes & "||(lcl_ext==""TIF"")"
%>

function validateFields() {
  var lcl_focus       = "";
  var lcl_false_count = 0;
  var isNumeric       = /^\d*$/;

  //Background Logo
		if (document.getElementById("logo_filenamebg").value!="") {
      lcl_logofilenamebg = document.getElementById("logo_filenamebg").value.toUpperCase();
      lcl_ext_start_pos  = lcl_logofilenamebg.indexOf(".");
      lcl_ext            = lcl_logofilenamebg.substr(lcl_ext_start_pos+1,lcl_logofilenamebg.length);

      if(<%=lcl_imgTypes%>) {
         clearMsg("findImageButtonbg");
     }else{
    					tabView.set('activeIndex',0);
         inlineMsg(document.getElementById("findImageButtonbg").id,'<strong>Invalid Value: </strong> The logo file extension is not valid. Valid file extensions:<br /><strong><%=lcl_imgTypesDisplay%></strong>',10,'findImageButtonbg');
         lcl_false_count = lcl_false_count + 1;
         lcl_focus       = document.getElementById("logofilename");
     }
  }else{
     clearMsg("findImageButtonbg");
  }

  //Logo
		if (document.getElementById("logo_filename").value!="") {
      lcl_logofilename = document.getElementById("logo_filename").value.toUpperCase();
      lcl_ext_start_pos = lcl_logofilename.indexOf(".");
      lcl_ext           = lcl_logofilename.substr(lcl_ext_start_pos+1,lcl_logofilename.length);

      if(<%=lcl_imgTypes%>) {
         clearMsg("findImageButton");
     }else{
    					tabView.set('activeIndex',0);
         inlineMsg(document.getElementById("findImageButton").id,'<strong>Invalid Value: </strong> The logo file extension is not valid. Valid file extensions:<br /><strong><%=lcl_imgTypesDisplay%></strong>',10,'findImageButton');
         lcl_false_count = lcl_false_count + 1;
         lcl_focus       = document.getElementById("logofilename");
     }
  }else{
     clearMsg("findImageButton");
  }

  //Website Size - Custom
  if(document.getElementById("website_size").value=="C") {
     if(document.getElementById("website_size_customsize").value=="") {
   					tabView.set('activeIndex',0);
        lcl_focus = document.getElementById("website_size_customsize");
        inlineMsg(document.getElementById("website_size_customsize").id,'<strong>Required Field Missing: </strong> Website Size (Custom Pixel Size)',10,'website_size_customsize');
        lcl_false_count = lcl_false_count + 1;
     }else{
        var rege = /^\d+$/;
        var Ok   = rege.exec(document.getElementById("website_size_customsize").value);

     			if (! Ok) {
       					tabView.set('activeIndex',0);
            inlineMsg(document.getElementById("website_size_customsize").id,'<strong>Invalid Value: </strong> Website Size (Custom Pixel Size) must be numeric.',10,'website_size_customsize');
            lcl_false_count = lcl_false_count + 1;
            lcl_focus       = document.getElementById("website_size_customsize");
        }else{
            clearMsg("website_size_customsize");
        }
     }
  }else{
     clearMsg("website_size_customsize");
  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     lcl_false_count = 0;
  }

  //Community Link Options
  lcl_totalCLRows = document.getElementById("totalCLRows").value;

  for (i=1; i<=lcl_totalCLRows; ++ i) {

     lcl_showSection_CL    = document.getElementById("showSection_CL_" + i).checked;
     lcl_showSection_SAVVY = document.getElementById("showSection_SAVVY_" + i).checked;

     //Display Order
     if(lcl_showSection_CL) {
        if(document.getElementById("displayorder_"+i).value != "") {
          	var Ok = isNumeric.test(document.getElementById("displayorder_"+i).value);
          	if(! Ok)	{
         					tabView.set('activeIndex',4);
          			 inlineMsg(document.getElementById("displayorder_"+i).id,'<strong>Invalid Value: </strong> The \"Display Order\" must be in a number format.',8,'displayorder_'+i);
              lcl_false_count = lcl_false_count + 1;

              if(lcl_false_count == 1) {
                 if(lcl_focus == "") {
                    lcl_focus = document.getElementById("displayorder_"+i);
                 }else{
                    lcl_focus = lcl_focus;
                 }
              }
           }else{
              clearMsg('displayorder_'+i);
         		}
        }
     }

     //# List Items - Savvy/IFRAME
     if(lcl_showSection_SAVVY) {
        if(document.getElementById("numListItemsShown_SAVVY_"+i).value != "") {
          	var Ok = isNumeric.test(document.getElementById("numListItemsShown_SAVVY_"+i).value);
          	if(! Ok)	{
         					tabView.set('activeIndex',4);
          			 inlineMsg(document.getElementById("numListItemsShown_SAVVY_"+i).id,'<strong>Invalid Value: </strong> The \"# List Items\" must be in a number format.',8,'numListItemsShown_SAVVY_'+i);
              lcl_false_count = lcl_false_count + 1;

              if(lcl_false_count == 1) {
                 if(lcl_focus == "") {
                    lcl_focus = document.getElementById("numListItemsShown_SAVVY_"+i);
                 }else{
                    lcl_focus = lcl_focus;
                 }
              }
           }else{
              clearMsg('numListItemsShown_SAVVY_'+i);
        		}
        }
     }

     //# List Items - CommunityLink
     if(document.getElementById("numListItemsShown_CL_"+i).value != "" && lcl_showSection_CL) {
       	var Ok = isNumeric.test(document.getElementById("numListItemsShown_CL_"+i).value);
       	if(! Ok)	{
      					tabView.set('activeIndex',4);
       			 inlineMsg(document.getElementById("numListItemsShown_CL_"+i).id,'<strong>Invalid Value: </strong> The \"# List Items\" must be in a number format.',8,'numListItemsShown_CL_'+i);
           lcl_false_count = lcl_false_count + 1;

           if(lcl_false_count == 1) {
              if(lcl_focus == "") {
                 lcl_focus = document.getElementById("numListItemsShown_CL_"+i);
              }else{
                 lcl_focus = lcl_focus;
              }
           }
        }else{
           clearMsg('numListItemsShown_CL_'+i);
      		}
     }
  }

  if(lcl_false_count > 0) {
     lcl_focus.focus();
     return false;
  }else{
     document.getElementById("communitylink_maint").submit();
     return true;
  }
}

function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL, p_returnAsHTMLLink, p_returnOnlyFileName) {
  w = 600;
  h = 400;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_showFolderStart = "";
  lcl_folderStart     = 0;

  //Determine which options will be displayed
  if((p_displayDocuments=="")||(p_displayDocuments==undefined)) {
      lcl_displayDocuments = "";
  }else{
      lcl_displayDocuments = "&displayDocuments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayActionLine=="")||(p_displayActionLine==undefined)) {
      lcl_displayActionLine = "";
  }else{
      lcl_displayActionLine = "&displayActionLine=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayPayments=="")||(p_displayPayments==undefined)) {
      lcl_displayPayments = "";
  }else{
      lcl_displayPayments = "&displayPayments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayURL=="")||(p_displayURL==undefined)) {
      lcl_displayURL = "";
  }else{
      lcl_displayURL = "&displayURL=Y";
  }

  if((p_returnAsHTMLLink=="")||(p_returnAsHTMLLink==undefined)) {
      lcl_returnAsHTMLLink = "";
  }else{
      lcl_returnAsHTMLLink = "&returnAsHTMLLink=" + p_returnAsHTMLLink;
  }

  if((p_returnOnlyFileName=="")||(p_returnOnlyFileName==undefined)) {
      lcl_returnOnlyFileName = "";
  }else{
      lcl_returnOnlyFileName = "&returnOnlyFileName=" + p_returnOnlyFileName;
  }

  if(lcl_folderStart > 0) {
     lcl_showFolderStart = "&folderStart=unpublished_documents";
  }

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += lcl_showFolderStart;
  pickerURL += lcl_displayDocuments;
  pickerURL += lcl_displayActionLine;
  pickerURL += lcl_displayPayments;
  pickerURL += lcl_displayURL;
  pickerURL += lcl_returnAsHTMLLink;
  pickerURL += lcl_returnOnlyFileName;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

function storeCaret (textEl) {
  if (textEl.createTextRange)
      textEl.caretPos = document.selection.createRange().duplicate();
}

function insertAtCaret (textEl, text) {
  if (textEl.createTextRange && textEl.caretPos) {
      var caretPos = textEl.caretPos;
      caretPos.text =
      caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
      text + ' ' : text;
  }
   else
      textEl.value = textEl.value + text;
}

function openWin(p_page, p_field_id, p_wintype, p_width, p_height) {
  if ((p_wintype=="")||(p_wintype==undefined)) {
       s_wintype="_picker";
  }else{
       s_wintype=p_wintype;
  }
  if ((p_width=="")||(p_width==undefined)) {
       s_width=600;
  }else{
       s_width=p_width;
  }
  if ((p_height=="")||(p_height==undefined)) {
       s_height=470;
  }else{
       s_height=p_height;
  }

  s_left = (screen.availWidth/2) - (s_width/2);
  s_top  = (screen.availHeight/2) - (s_height/2);

  if((p_field_id=="")||(p_field_id!=undefined)) {
      lcl_url = p_page;
  }else{
      lcl_url = p_page + "?fieldid=" + p_field_id;
  }

		eval('window.open("' + lcl_url + '", "' + s_wintype + '", "width=' + s_width + ',height=' + s_height + ',left=' + s_left + ',top=' + s_top + 'status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes")');
}

function changePreviewColor(iFieldID) {
  lcl_color = document.getElementById(iFieldID).value;

		var rege = /^[0-9a-f]{3,6}$/i;
		var Ok = rege.exec(lcl_color);

  if ( Ok ) {
      $('#' + iFieldID + '_previewcolor').css('background-color',lcl_color);
      //document.getElementById(iFieldID + '_previewcolor').style.backgroundColor=lcl_color;
  }
}

function resetFields(iRowCount) {
<%
 'Get the defaults
  lcl_fontcolor_default                   = "000000"
  lcl_sectionheader_bgcolor_default       = "ffffff"
  lcl_sectionheader_linecolor_default     = "000000"
  lcl_sectiontext_bgcolor_default         = "ffffff"
  lcl_sectionheader_fonttype_default      = getCLOptionDefault("SECTIONHEADER_FONTTYPE")
  lcl_sectiontext_fonttype_default        = getCLOptionDefault("SECTIONTEXT_FONTTYPE")
  lcl_sectionlinks_alignment_default      = getCLOptionDefault("SECTIONLINKS_ALIGN")
  lcl_sectionlinks_fonttype_default       = getCLOptionDefault("SECTIONLINKS_FONTTYPE")
  lcl_sectionlinks_fontcolor_default      = "800000"
  lcl_sectionlinks_fontcolorhover_default = "800000"
  lcl_portalcolumn_default                = getCLOptionDefault("PORTALCOLUMNS")
  lcl_displayorder_default                = "1"
%>
  document.getElementById("featurename_" + iRowCount).value = document.getElementById("featurename_original_" + iRowCount).value;

  lcl_showSection_CL    = document.getElementById("showSection_CL_" + iRowCount).checked;
  lcl_showSection_SAVVY = document.getElementById("showSection_SAVVY_" + iRowCount).checked;

  //CommunityLink options
  if(lcl_showSection_CL) {
     document.getElementById("sectionheader_bgcolor_CL_"       + iRowCount).value = "<%=lcl_sectionheader_bgcolor_default%>";
     document.getElementById("sectionheader_linecolor_CL_"     + iRowCount).value = "<%=lcl_sectionheader_linecolor_default%>";
     document.getElementById("sectionheader_fonttype_CL_"      + iRowCount).value = "<%=lcl_sectionheader_fonttype_default%>";
     document.getElementById("sectionheader_fontcolor_CL_"     + iRowCount).value = "<%=lcl_fontcolor_default%>"
     document.getElementById("sectiontext_bgcolor_CL_"         + iRowCount).value = "<%=lcl_sectiontext_bgcolor_default%>";
     document.getElementById("sectiontext_fonttype_CL_"        + iRowCount).value = "<%=lcl_sectiontext_fonttype_default%>";
     document.getElementById("sectiontext_fontcolor_CL_"       + iRowCount).value = "<%=lcl_fontcolor_default%>";
     document.getElementById("sectionlinks_alignment_CL_"      + iRowCount).value = "<%=lcl_sectionlinks_alignment_default%>";
     document.getElementById("sectionlinks_fonttype_CL_"       + iRowCount).value = "<%=lcl_sectionlinks_fonttype_default%>";
     document.getElementById("sectionlinks_fontcolor_CL_"      + iRowCount).value = "<%=lcl_sectionlinks_fontcolor_default%>";
     document.getElementById("sectionlinks_fontcolorhover_CL_" + iRowCount).value = "<%=lcl_sectionlinks_fontcolorhover_default%>";
     document.getElementById("numListItemsShown_CL_"           + iRowCount).value = document.getElementById("numListItemsShown_original_" + iRowCount).value;
     document.getElementById("portalcolumn_"                   + iRowCount).value = "<%=lcl_portalcolumn_default%>";
     document.getElementById("displayorder_"                   + iRowCount).value = "<%=lcl_displayorder_default%>";

     changePreviewColor("sectionheader_bgcolor_CL_"   + iRowCount);
     changePreviewColor("sectionheader_linecolor_CL_" + iRowCount);
     changePreviewColor("sectionheader_fontcolor_CL_" + iRowCount);
     changePreviewColor("sectiontext_bgcolor_CL_"     + iRowCount);
     changePreviewColor("sectiontext_fontcolor_CL_"   + iRowCount);
     changePreviewColor("sectionlinks_fontcolor_CL_"  + iRowCount);
  }

  //Savvy/IFRAME options
  if(lcl_showSection_SAVVY) {
     document.getElementById("sectionheader_bgcolor_SAVVY_"       + iRowCount).value = "<%=lcl_sectionheader_bgcolor_default%>";
     document.getElementById("sectionheader_linecolor_SAVVY_"     + iRowCount).value = "<%=lcl_sectionheader_linecolor_default%>";
     document.getElementById("sectionheader_fonttype_SAVVY_"      + iRowCount).value = "<%=lcl_sectionheader_fonttype_default%>";
     document.getElementById("sectionheader_fontcolor_SAVVY_"     + iRowCount).value = "<%=lcl_fontcolor_default%>";
     document.getElementById("sectiontext_bgcolor_SAVVY_"         + iRowCount).value = "<%=lcl_sectiontext_bgcolor_default%>";
     document.getElementById("sectiontext_fonttype_SAVVY_"        + iRowCount).value = "<%=lcl_sectiontext_fonttype_default%>";
     document.getElementById("sectiontext_fontcolor_SAVVY_"       + iRowCount).value = "<%=lcl_fontcolor_default%>";
     document.getElementById("sectionlinks_alignment_SAVVY_"      + iRowCount).value = "<%=lcl_sectionlinks_alignment_default%>";
     document.getElementById("sectionlinks_fonttype_SAVVY_"       + iRowCount).value = "<%=lcl_sectionlinks_fonttype_default%>";
     document.getElementById("sectionlinks_fontcolor_SAVVY_"      + iRowCount).value = "<%=lcl_sectionlinks_fontcolor_default%>";
     document.getElementById("sectionlinks_fontcolorhover_SAVVY_" + iRowCount).value = "<%=lcl_sectionlinks_fontcolorhover_default%>";

     document.getElementById("numListItemsShown_SAVVY_"       + iRowCount).value = document.getElementById("numListItemsShown_original_" + iRowCount).value;

     changePreviewColor("sectionheader_bgcolor_SAVVY_"   + iRowCount);
     changePreviewColor("sectionheader_linecolor_SAVVY_" + iRowCount);
     changePreviewColor("sectionheader_fontcolor_SAVVY_" + iRowCount);
     changePreviewColor("sectiontext_bgcolor_SAVVY_"     + iRowCount);
     changePreviewColor("sectiontext_fontcolor_SAVVY_"   + iRowCount);
     changePreviewColor("sectionlinks_fontcolor_SAVVY_"  + iRowCount);
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

function changeElementStyles(p_viewLinkID, p_fontsize, p_fonttype, p_fontcolor, p_underline, p_backgroundcolor) {

  if(p_viewLinkID!="") {
     var lcl_viewlink = document.getElementById(p_viewLinkID);

     if(p_fontsize!="") {
        lcl_viewlink.style.fontSize = p_fontsize+'px';
     }

     if(p_fonttype!="") {
        lcl_viewlink.style.fontFamily = p_fonttype;
     }

     if(p_fontcolor!="") {
        lcl_viewlink.style.color = '#' + p_fontcolor;
     }

     if(p_underline!="") {
        lcl_viewlink.style.textDecoration = p_underline;
     }

     if(p_backgroundcolor!="") {
        lcl_viewlink.style.backgroundColor = '#' + p_backgroundcolor;
     }
  }
}

function enableDisableViewALLURL(iRowID, iIsOnChange)
{
   var lcl_value      = $('#sectionlinks_viewall_urltype_' + iRowID).val();
   var lcl_isOnChange = 'N';

   if(iIsOnChange != '')
   {
      lcl_isOnChange = iIsOnChange;
   }

   $('#sectionlinks_viewall_url_' + iRowID).prop('disabled',true);

   if(lcl_value == 'custom')
   {
      $('#sectionlinks_viewall_url_' + iRowID).prop('disabled',false);

      if(lcl_isOnChange == 'Y')
      {
         $('#sectionlinks_viewall_url_' + iRowID).focus();
      }
   }
}

window.onload = function(){
  //Set the formname to the following: "document.FORM_NAME.  (NO quotes or double-quotes)!
//  formName = document.communitylink_maint;

 	//Create our helper object that will show the item while dragging
//	 dragHelper = document.createElement('DIV');
//	 dragHelper.style.cssText = 'position:absolute;display:none;';

  //Identify the number of DragContainers you have in your screen.  These are array values.
//	 CreateDragContainer(
//  		document.getElementById('DragContainer1'),
//		  document.getElementById('DragContainer2')
  		//document.getElementById('DragContainer3')
// 	);

//	 document.body.appendChild(dragHelper);

  <%=lcl_onload%>
}
//-->
</script>
</head>
<body>
<!--<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<% 'lcl_onload%>"> -->

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  lcl_labelcolumn_width = "150"

  response.write "<div id=""content"">" & vbcrlf
  'response.write "  <div id=""centercontent"">" & vbcrlf

  response.write "<form name=""communitylink_maint"" id=""communitylink_maint"" action=""communitylink_action.asp"" method=""post"">" & vbcrlf
  response.write "  <input type=""hidden"" name=""communitylinkid"" id=""communitylinkid"" value=""" & lcl_communitylinkid & """ size=""5"" maxlength=""10"" />" & vbcrlf

 'BEGIN: Page Title and Screen Messages ---------------------------------------
  response.write "<div style=""margin-bottom:10px;"">" & vbcrlf
  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width:1000px"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td><font size=""+1""><strong>" & session("sOrgName") & "&nbsp;" & lcl_pagetitle & "</strong></font></td>" & vbcrlf
  response.write "      <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

                  displayButtons "MAINT"

  response.write "</div>" & vbcrlf
 'END: Page Title and Screen Messages -----------------------------------------

 'BEGIN: Layout Options -------------------------------------------------------
  lcl_checked_egovhome = ""

  if lcl_isEgovHomePage then
     lcl_checked_egovhome = " checked=""checked"""
  end if

  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Layout Options</legend>" & vbcrlf
  response.write "<table id=""tableLayoutOptions"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"" colspan=""2"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

  response.write "            <tr>" & vbcrlf
  response.write "                <td colspan=""2"">" & vbcrlf
  response.write "                    Make CommunityLink E-Gov Home Page:" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""isEgovHomePage"" id=""isEgovHomePage"" value=""on""" & lcl_checked_egovhome & " />" & vbcrlf
  response.write "                    <br /><br />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Website Size
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Size:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""website_size"" id=""website_size"" onchange=""setupCustomSize();"">" & vbcrlf
                                        displayCommunityLinkOptions "WEBSITE_SIZE", lcl_website_size
  response.write "                    </select>&nbsp;" & vbcrlf
  response.write "                    <span id=""website_size_customsize_span"">" & vbcrlf
  response.write "                    <input type=""text"" name=""website_size_customsize"" id=""website_size_customsize"" value=""" & lcl_website_size_customsize & """ size=""5"" maxlength=""10"" onchange=""clearMsg('website_size_customsize');"" />&nbsp;" & vbcrlf
  response.write "                    <font style=""font-size:10px; color:#800000"">(All sizes in pixels)</font></span>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Website Alignment
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Alignment:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""website_alignment"" id=""website_alignment"">" & vbcrlf
                                        displayCommunityLinkOptions "WEBSITE_ALIGN", lcl_website_alignment
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Website Background Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Background Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "website_bgcolor", lcl_website_bgcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('website_bgcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  response.write "            <tr><td colspan=""3"">&nbsp;</td></tr>" & vbcrlf

 'Show Logo
  if lcl_showlogo then
     lcl_checked_logo = " checked=""checked"""
  else
     lcl_checked_logo = ""
  end if

  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Logo:</td>" & vbcrlf
  response.write "                <td><input type=""checkbox"" name=""showlogo"" id=""showlogo"" value=""on""" & lcl_checked_logo & " /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Logo Alignment
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Alignment:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""logo_alignment"" id=""logo_alignment"">" & vbcrlf
                                        displayCommunityLinkOptions "WEBSITE_LOGO_ALIGN", lcl_logo_alignment
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Logo Filename
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Logo:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <input type=""input"" name=""logo_filename"" id=""logo_filename"" value=""" & lcl_logo_filename & """ size=""50"" maxlength=""500"" onchange=""clearMsg('findImageButton');"" />&nbsp;" & vbcrlf
  response.write "                    <input type=""button"" name=""findImageButton"" id=""findImageButton"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButton');doPicker('communitylink_maint.logo_filename','Y','','','','N','Y');"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Logo Filename - Background
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Background Logo:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <input type=""input"" name=""logo_filenamebg"" id=""logo_filenamebg"" value=""" & lcl_logo_filenamebg & """ size=""50"" maxlength=""500"" onchange=""clearMsg('findImageButtonbg');"" />&nbsp;" & vbcrlf
  response.write "                    <input type=""button"" name=""findImageButtonbg"" id=""findImageButtonbg"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButtonbg');doPicker('communitylink_maint.logo_filenamebg','Y','','','','N','Y');"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
 'END: Layout Options ---------------------------------------------------------

 'BEGIN: Top Bar Options ------------------------------------------------------
  lcl_checked_topbar = ""

  if lcl_showtopbar then
     lcl_checked_topbar = " checked=""checked"""
  end if

  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Top Bar Options</legend>" & vbcrlf

  response.write "<table>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Top Bar
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Top Bar:</td>" & vbcrlf
  response.write "                <td><input type=""checkbox"" name=""showtopbar"" id=""showtopbar"" value=""on""" & lcl_checked_topbar & " /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Background Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Background Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "topbar_bgcolor", lcl_topbar_bgcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('topbar_bgcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Font Type
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Type:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""topbar_fonttype"" id=""topbar_fonttype"">" & vbcrlf
                                        displayCommunityLinkOptions "TOPBAR_FONTTYPE", lcl_topbar_fonttype
  response.write "                    </select>" & vbcrf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Font Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "topbar_fontcolor", lcl_topbar_fontcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('topbar_fontcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Font Color - Hover
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Color<br />(mouseover):</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "topbar_fontcolorhover", lcl_topbar_fontcolorhover, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('topbar_fontcolorhover');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf


  response.write "</fieldset>" & vbcrlf
 'END: Top Bar Options --------------------------------------------------------

 'BEGIN: Page Header Options --------------------------------------------------
  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Page Header Options</legend>" & vbcrlf

  response.write "<table>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Page Header
  if lcl_showpageheader then
     lcl_checked_pageheader = " checked=""checked"""
  else
     lcl_checked_pageheader = ""
  end if

  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Page Header:</td>" & vbcrlf
  response.write "                <td><input type=""checkbox"" name=""showpageheader"" id=""showpageheader"" value=""on""" & lcl_checked_pageheader & " /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Background Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Background Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "pageheader_bgcolor", lcl_pageheader_bgcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('pageheader_bgcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Page Header Alignment
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Alignment:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""pageheader_alignment"" id=""pageheader_alignment"">" & vbcrlf
                                        displayCommunityLinkOptions "PAGEHEADER_ALIGN", lcl_pageheader_alignment
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Page Font Size
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Size:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <input type=""text"" name=""pageheader_fontsize"" id=""pageheader_fontsize"" value=""" & lcl_pageheader_fontsize & """ size=""3"" maxlength=""3"" /> <em style=""font-size:10px; color:#ff0000;"">(in pixels)</em>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Page Header Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "pageheader_fontcolor", lcl_pageheader_fontcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('pageheader_fontcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Page Header - Font Type
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Type:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""pageheader_fonttype"" id=""pageheader_fonttype"">" & vbcrlf
                                        displayCommunityLinkOptions "PAGEHEADER_FONTTYPE", lcl_pageheader_fonttype
  response.write "                    </select>" & vbcrf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  response.write "            <tr><td colspan=""3"">&nbsp;</td></tr>" & vbcrlf

 'Show RSS
  if lcl_showRSS then
     lcl_checked_showRSS = " checked=""checked"""
  else
     lcl_checked_showRSS = ""
  end if

  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show RSS (icon):</td>" & vbcrlf
  response.write "                <td colspan=""2"">" & vbcrlf
  response.write "                    <input type=""checkbox"" name=""showRSS"" id=""showRSS"" value=""on""" & lcl_checked_showRSS & " />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'URL - Twitter
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - Twitter:</td>" & vbcrlf
  response.write "                <td colspan=""2"">" & vbcrlf
  response.write "                    <input type=""input"" name=""url_twitter"" id=""url_twitter"" value=""" & lcl_url_twitter & """ size=""80"" maxlength=""500"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'URL - Facebook
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - Facebook:</td>" & vbcrlf
  response.write "                <td colspan=""2"">" & vbcrlf
  response.write "                    <input type=""input"" name=""url_facebook"" id=""url_facebook"" value=""" & lcl_url_facebook & """ size=""80"" maxlength=""500"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'URL - MySpace
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - MySpace:</td>" & vbcrlf
  response.write "                <td colspan=""2"">" & vbcrlf
  response.write "                    <input type=""input"" name=""url_myspace"" id=""url_myspace"" value=""" & lcl_url_myspace & """ size=""80"" maxlength=""500"" />" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'URL - Blogger
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - Blogger:</td>" & vbcrlf
  response.write "                <td colspan=""2"">" & vbcrlf
  response.write "                    <input type=""input"" name=""url_blogger"" id=""url_blogger"" value=""" & lcl_url_blogger & """ size=""80"" maxlength=""500"" />&nbsp;" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  response.write "</fieldset>" & vbcrlf
 'END: Page Header Options ----------------------------------------------------

 'BEGIN: Side Bar Options -----------------------------------------------------
  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Side Bar Options</legend>" & vbcrlf

  response.write "<table>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Side Menubar
  if lcl_showsidemenubar then
     lcl_checked_sidemenubar = " checked=""checked"""
  else
     lcl_checked_sidemenubar = ""
  end if

  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Side Menu Bar:</td>" & vbcrlf
  response.write "                <td><input type=""checkbox"" name=""showsidemenubar"" id=""showsidemenubar"" value=""on""" & lcl_checked_sidemenubar & " /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Side Menubar Alignment
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Alignment:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""sidemenubar_alignment"" id=""sidemenubar_alignment"">" & vbcrlf
                                        displayCommunityLinkOptions "SIDEMENUBAR_ALIGN", lcl_sidemenubar_alignment
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Side Menubar Option Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Option Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "sidemenuoption_bgcolor", lcl_sidemenuoption_bgcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_bgcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Side Menubar Option Color - Hover
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Option Color<br />(mouseover):</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "sidemenuoption_bgcolorhover", lcl_sidemenuoption_bgcolorhover, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_bgcolorhover');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf

 'Side Menubar Option Alignment
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Text Alignment:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""sidemenuoption_alignment"" id=""sidemenuoption_alignment"">" & vbcrlf
                                        displayCommunityLinkOptions "SIDEMENUOPT_TEXTALIGN", lcl_sidemenuoption_alignment
  response.write "                    </select>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Side Menubar Option - Font Type
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Option Font Type:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""sidemenuoption_fonttype"" id=""sidemenuoption_fonttype"">" & vbcrlf
                                        displayCommunityLinkOptions "SIDEMENUOPT_FONTTYPE", lcl_sidemenuoption_fonttype
  response.write "                    </select>" & vbcrf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Side Menubar Option - Font Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Option Font Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "sidemenuoption_fontcolor", lcl_sidemenuoption_fontcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_fontcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Side Menubar Option - Font Color - Hover
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Option Font Color<br />(mouseover):</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "sidemenuoption_fontcolorhover", lcl_sidemenuoption_fontcolorhover, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_fontcolorhover');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  response.write "</fieldset>" & vbcrlf
 'END: Side Bar Options -------------------------------------------------------

 'BEGIN: Footer Options -------------------------------------------------------
  lcl_checked_footer = ""

  if lcl_showfooter then
     lcl_checked_footer = " checked=""checked"""
  end if

  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Footer Options</legend>" & vbcrlf

  response.write "<table>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Footer:</td>" & vbcrlf
  response.write "                <td><input type=""checkbox"" name=""showfooter"" id=""showfooter"" value=""on""" & lcl_checked_footer & " /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Background Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Background Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "footer_bgcolor", lcl_footer_bgcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('footer_bgcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Font Type
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Type:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
  response.write "                    <select name=""footer_fonttype"" id=""footer_fonttype"">" & vbcrlf
                                        displayCommunityLinkOptions "FOOTER_FONTTYPE", lcl_footer_fonttype
  response.write "                    </select>" & vbcrf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Font Color
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Color:</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "footer_fontcolor", lcl_footer_fontcolor, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('footer_fontcolor');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf

 'Font Color - Hover
  response.write "            <tr>" & vbcrlf
  response.write "                <td>Font Color<br />(mouseover):</td>" & vbcrlf
  response.write "                <td>" & vbcrlf
                                      setupColorSelection "footer_fontcolorhover", lcl_footer_fontcolorhover, 1
                                      lcl_scripts = lcl_scripts & "changePreviewColor('footer_fontcolorhover');" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  response.write "</fieldset>" & vbcrlf
 'END: Footer Options ---------------------------------------------------------

 'BEGIN: CommunityLink Options ------------------------------------------------
  iBGColor                    = "#eeeeee"
  iRowCount                   = 0
  lcl_features_shown          = ""
  iTotalCLFeaturesAvail       = getCLFeatAvailCount(session("orgid"))
  lcl_isCommunityLinkOn       = 0
  lcl_isSavvyOn               = 0
  lcl_showsectionborder       = 0
  lcl_sectionbordercolor      = "000000"
  lcl_sectionheader_bgcolor   = "ffffff"
  lcl_sectionheader_linecolor = "000000"
  lcl_sectionheader_fonttype  = getCLOptionDefault("SECTIONHEADER_FONTTYPE")
  lcl_sectionheader_fontcolor = "000000"
  lcl_sectionheader_fontsize  = "11"
  lcl_sectionheader_isbold    = 1
  lcl_sectionheader_isitalic  = 0
  lcl_sectiontext_bgcolor     = "ffffff"
  lcl_sectiontext_fonttype    = getCLOptionDefault("SECTIONTEXT_FONTTYPE")
  lcl_sectiontext_fontcolor   = "000000"
  lcl_sectionlinks_alignment  = getCLOptionDefault("SECTIONLINKS_ALIGN")
  lcl_sectionlinks_fonttype   = getCLOptionDefault("SECTIONLINKS_FONTTYPE")
  lcl_sectionlinks_fontcolor  = "800000"
  lcl_viewall_urltype         = "default"
  lcl_viewall_url_wintype     = "samewindow"

  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>CommunityLink Options</legend>" & vbcrlf

  response.write "<table>" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"" class=""communitylink_table"">" & vbcrlf
  response.write "            <tr>" & vbcrlf
  response.write "                <th align=""left"">CommunityLink Sections</th>" & vbcrlf
  response.write "                <th nowrap=""nowrap"">Community Link</th>" & vbcrlf
  response.write "                <th nowrap=""nowrap"">Savvy/IFRAME</th>" & vbcrlf
  response.write "            </tr>" & vbcrlf

  sSQL = " SELECT f.featureid, "
  sSQL = sSQL & " isnull(f.CL_portaltype,'') AS portaltype, "
  sSQL = sSQL & " f.CommunityLinkOn, "
  sSQL = sSQL & " isnull(cl.featurename, isnull(otf.featurename, f.featurename)) AS featurename, "
  sSQL = sSQL & " isnull(otf.featurename, f.featurename) AS featurename_original, "
  sSQL = sSQL & " isnull(cl.portalcolumn, 1) AS portalcolumn, "
  sSQL = sSQL & " isnull(cl.displayorder, 1) AS displayorder, "
  sSQL = sSQL & " cl.rss_feedid, "
  sSQL = sSQL & " isnull(cl.numListItemsShown_CL, f.CL_numListItems) AS numListItemsShown_CL, "
  sSQL = sSQL & " isnull(cl.numListItemsShown_SAVVY, f.CL_numListItems) AS numListItemsShown_SAVVY, "
  sSQL = sSQL & " f.CL_numListItems AS numListItemsShown_original, "
  sSQL = sSQL & " isnull(cl.isCommunityLinkOn,"                  & lcl_isCommunityLinkOn       & ") AS isCommunityLinkOn, "
  sSQL = sSQL & " isnull(cl.isSavvyOn,"                          & lcl_isSavvyOn               & ") AS isSavvyOn, "
  sSQL = sSQL & " isnull(cl.showsectionborder_CL,"               & lcl_showsectionborder       & ") AS showsectionborder_CL, "
  sSQL = sSQL & " isnull(cl.showsectionborder_SAVVY,"            & lcl_showsectionborder       & ") AS showsectionborder_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionbordercolor_CL,'"             & lcl_sectionbordercolor      & "') AS sectionbordercolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionbordercolor_SAVVY,'"          & lcl_sectionbordercolor      & "') AS sectionbordercolor_SAVVY, "
  sSQL = sSQL & " cl.sectionbackgroundcolor_CL, "
  sSQL = sSQL & " cl.sectionbackgroundcolor_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionheader_bgcolor_CL,'"          & lcl_sectionheader_bgcolor   & "') AS sectionheader_bgcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_bgcolor_SAVVY,'"       & lcl_sectionheader_bgcolor   & "') AS sectionheader_bgcolor_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionheader_linecolor_CL,'"        & lcl_sectionheader_linecolor & "') AS sectionheader_linecolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_linecolor_SAVVY,'"     & lcl_sectionheader_linecolor & "') AS sectionheader_linecolor_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionheader_fonttype_CL,'"         & lcl_sectionheader_fonttype  & "') AS sectionheader_fonttype_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_fonttype_SAVVY,'"      & lcl_sectionheader_fonttype  & "') AS sectionheader_fonttype_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionheader_fontcolor_CL,'"        & lcl_sectionheader_fontcolor & "') AS sectionheader_fontcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_fontcolor_SAVVY,'"     & lcl_sectionheader_fontcolor & "') AS sectionheader_fontcolor_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionheader_fontsize_CL,'"         & lcl_sectionheader_fontsize  & "') AS sectionheader_fontsize_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_fontsize_SAVVY,'"      & lcl_sectionheader_fontsize  & "') AS sectionheader_fontsize_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionheader_isbold_CL,'"           & lcl_sectionheader_isbold    & "') AS sectionheader_isbold_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_isbold_SAVVY,'"        & lcl_sectionheader_isbold    & "') AS sectionheader_isbold_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionheader_isitalic_CL,'"         & lcl_sectionheader_isitalic  & "') AS sectionheader_isitalic_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_isitalic_SAVVY,'"      & lcl_sectionheader_isitalic  & "') AS sectionheader_isitalic_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectiontext_bgcolor_CL,'"            & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_bgcolor_SAVVY,'"         & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolor_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectiontext_bgcolorhover_CL,'"       & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolorhover_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_bgcolorhover_SAVVY,'"    & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolorhover_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectiontext_fonttype_CL,'"           & lcl_sectiontext_fonttype    & "') AS sectiontext_fonttype_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_fonttype_SAVVY,'"        & lcl_sectiontext_fonttype    & "') AS sectiontext_fonttype_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontcolor_CL,'"          & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontcolor_SAVVY,'"       & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolor_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontcolorhover_CL,'"     & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolorhover_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontcolorhover_SAVVY,'"  & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolorhover_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontsize_CL,'"           & lcl_sectiontext_fontsize    & "') AS sectiontext_fontsize_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontsize_SAVVY,'"        & lcl_sectiontext_fontsize    & "') AS sectiontext_fontsize_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionlinks_alignment_CL,'"         & lcl_sectionlinks_alignment  & "') AS sectionlinks_alignment_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_alignment_SAVVY,'"      & lcl_sectionlinks_alignment  & "') AS sectionlinks_alignment_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fonttype_CL,'"          & lcl_sectionlinks_fonttype   & "') AS sectionlinks_fonttype_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fonttype_SAVVY,'"       & lcl_sectionlinks_fonttype   & "') AS sectionlinks_fonttype_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolor_CL,'"         & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolor_SAVVY,'"      & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolor_SAVVY, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolorhover_CL,'"    & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolorhover_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolorhover_SAVVY,'" & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolorhover_SAVVY, "
  sSQL = sSQL & " isnull(cl.viewall_urltype_CL, '"               & lcl_viewall_urltype         & "') AS viewall_urltype_CL, "
  sSQL = sSQL & " isnull(cl.viewall_urltype_SAVVY, '"            & lcl_viewall_urltype         & "') AS viewall_urltype_SAVVY, "
  sSQL = sSQL & " cl.viewall_url_CL, "
  sSQL = sSQL & " cl.viewall_url_SAVVY, "
  sSQL = sSQL & " isnull(cl.viewall_url_wintype_CL, '"           & lcl_viewall_url_wintype     & "') AS viewall_url_wintype_CL, "
  sSQL = sSQL & " isnull(cl.viewall_url_wintype_SAVVY, '"        & lcl_viewall_url_wintype     & "') AS viewall_url_wintype_SAVVY, "
  sSQL = sSQL & " query_filter "
  sSQL = sSQL & " FROM egov_communitylink_displayorgfeatures cl "
  sSQL = sSQL &      " RIGHT OUTER JOIN egov_organizations_to_features otf "
  sSQL = sSQL &      " INNER JOIN egov_organization_features f "
  sSQL = sSQL &      " ON otf.featureid = f.featureid "
  sSQL = sSQL &      " ON f.featureid = cl.featureid "
  sSQL = sSQL &      " AND cl.orgid = otf.orgid "
  sSQL = sSQL & " WHERE f.haspublicview = 1 "
  sSQL = sSQL & " AND f.CommunityLinkOn = 1 "
  sSQL = sSQL & " AND otf.orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY cl.isCommunityLinkOn DESC, "
  sSQL = sSQL &          " isnull(cl.portalcolumn, 0), "
  sSQL = sSQL &          " isnull(cl.displayorder, 0), "
  sSQL = sSQL &          " isnull(cl.featurename, isnull(otf.featurename, f.featurename)), "
  sSQL = sSQL &          " cl.isSavvyOn DESC "

  set oCLFeatures = Server.CreateObject("ADODB.Recordset")
  oCLFeatures.Open sSQL, Application("DSN"), 3, 1

  if not oCLFeatures.eof then
     do while not oCLFeatures.eof
        iBGColor  = changeBGColor(iBGColor,"#eeeeee","#ffffff")
        iRowCount = iRowCount + 1

        response.write "            <tr valign=""top"" bgcolor=""" & iBGColor & """>" & vbcrlf

        displayCommunityLinkFeatureOptions "CL", _
                                           iRowCount, _
                                           iTotalCLFeaturesAvail, _
                                           iBGColor, _
                                           lcl_scripts, _
                                           oCLFeatures("featureid"), _
                                           oCLFeatures("featurename"), _
                                           oCLFeatures("featurename_original"), _
                                           oCLFeatures("portalcolumn"), _
                                           oCLFeatures("displayorder"), _
                                           oCLFeatures("rss_feedid"), _
                                           oCLFeatures("numListItemsShown_CL"), _
                                           oCLFeatures("numListItemsShown_original"), _
                                           oCLFeatures("isCommunityLinkOn"), _
                                           oCLFeatures("showsectionborder_CL"), _
                                           oCLFeatures("sectionbordercolor_CL"), _
                                           oCLFeatures("sectionbackgroundcolor_CL"), _
                                           oCLFeatures("sectionheader_bgcolor_CL"), _
                                           oCLFeatures("sectionheader_linecolor_CL"), _
                                           oCLFeatures("sectionheader_fonttype_CL"), _
                                           oCLFeatures("sectionheader_fontcolor_CL"), _
                                           oCLFeatures("sectionheader_fontsize_CL"), _
                                           oCLFeatures("sectionheader_isbold_CL"), _
                                           oCLFeatures("sectionheader_isitalic_CL"), _
                                           oCLFeatures("sectiontext_bgcolor_CL"), _
                                           oCLFeatures("sectiontext_bgcolorhover_CL"), _
                                           oCLFeatures("sectiontext_fonttype_CL"), _
                                           oCLFeatures("sectiontext_fontcolor_CL"), _
                                           oCLFeatures("sectiontext_fontcolorhover_CL"), _
                                           oCLFeatures("sectiontext_fontsize_CL"), _
                                           oCLFeatures("sectionlinks_alignment_CL"), _
                                           oCLFeatures("sectionlinks_fonttype_CL"), _
                                           oCLFeatures("sectionlinks_fontcolor_CL"), _
                                           oCLFeatures("sectionlinks_fontcolorhover_CL"), _
                                           oCLFeatures("viewall_urltype_CL"), _
                                           oCLFeatures("viewall_url_CL"), _
                                           oCLFeatures("viewall_url_wintype_CL"), _
                                           oCLFeatures("query_filter")

        displayCommunityLinkFeatureOptions "SAVVY", _
                                           iRowCount, _
                                           iTotalCLFeaturesAvail, _
                                           iBGColor, _
                                           lcl_scripts, _
                                           oCLFeatures("featureid"), _
                                           oCLFeatures("featurename"), _
                                           oCLFeatures("featurename_original"), _
                                           oCLFeatures("portalcolumn"), _
                                           oCLFeatures("displayorder"), _
                                           oCLFeatures("rss_feedid"), _
                                           oCLFeatures("numListItemsShown_SAVVY"), _
                                           oCLFeatures("numListItemsShown_original"), _
                                           oCLFeatures("isSavvyOn"), _
                                           oCLFeatures("showsectionborder_SAVVY"), _
                                           oCLFeatures("sectionbordercolor_SAVVY"), _
                                           oCLFeatures("sectionbackgroundcolor_SAVVY"), _
                                           oCLFeatures("sectionheader_bgcolor_SAVVY"), _
                                           oCLFeatures("sectionheader_linecolor_SAVVY"), _
                                           oCLFeatures("sectionheader_fonttype_SAVVY"), _
                                           oCLFeatures("sectionheader_fontcolor_SAVVY"), _
                                           oCLFeatures("sectionheader_fontsize_SAVVY"), _
                                           oCLFeatures("sectionheader_isbold_SAVVY"), _
                                           oCLFeatures("sectionheader_isitalic_SAVVY"), _
                                           oCLFeatures("sectiontext_bgcolor_SAVVY"), _
                                           oCLFeatures("sectiontext_bgcolorhover_SAVVY"), _
                                           oCLFeatures("sectiontext_fonttype_SAVVY"), _
                                           oCLFeatures("sectiontext_fontcolor_SAVVY"), _
                                           oCLFeatures("sectiontext_fontcolorhover_SAVVY"), _
                                           oCLFeatures("sectiontext_fontsize_SAVVY"), _
                                           oCLFeatures("sectionlinks_alignment_SAVVY"), _
                                           oCLFeatures("sectionlinks_fonttype_SAVVY"), _
                                           oCLFeatures("sectionlinks_fontcolor_SAVVY"), _
                                           oCLFeatures("sectionlinks_fontcolorhover_SAVVY"), _
                                           oCLFeatures("viewall_urltype_SAVVY"), _
                                           oCLFeatures("viewall_url_SAVVY"), _
                                           oCLFeatures("viewall_url_wintype_SAVVY"), _
                                           oCLFeatures("query_filter")

        response.write "            </tr>" & vbcrlf

        oCLFeatures.movenext
     loop
  end if

  oCLFeatures.close
  set oCLFeatures = nothing

 'Total Rows
  response.write "            <tr>" & vbcrlf
  response.write "                <td><input type=""hidden"" name=""totalCLRows"" id=""totalCLRows"" value=""" & iRowCount & """ size=""3"" maxlength=""10"" /></td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
 'END: CommunityLink Options --------------------------------------------------

 'BEGIN: Tabs (headers) -------------------------------------------------------
  'response.write "<div id=""demo"" class=""yui-navset"" style=""width:1000px"">" & vbcrlf
  'response.write "  <ul class=""yui-nav"">" & vbcrlf
  'response.write "  		<li><a href=""#tab1""><em>Layout Options</em></a></li>" & vbcrlf
  'response.write "  		<li><a href=""#tab2""><em>Top Bar Options</em></a></li>" & vbcrlf
  'response.write "  		<li><a href=""#tab3""><em>Page Header Options</em></a></li>" & vbcrlf
  'response.write "  		<li><a href=""#tab4""><em>Side Bar Menu Options</em></a></li>" & vbcrlf
  'response.write "  		<li><a href=""#tab5""><em>Footer Options</em></a></li>" & vbcrlf
  'response.write "  		<li><a href=""#tab6""><em>CommunityLink Options</em></a></li>" & vbcrlf
  'response.write "  </ul>" & vbcrlf
  'response.write "<div class=""yui-content"">" & vbcrlf
 'END: Tabs (headers) ---------------------------------------------------------

 'BEGIN: Layout Options -------------------------------------------------------
'  response.write "<div id=""tab1"">" & vbcrlf
'  response.write "<table border=""0"" bordercolor=""#00000ff"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""800"" style=""margin-top:10px;"">" & vbcrlf
'  response.write "  <tr>" & vbcrlf
'  response.write "      <td valign=""top"" colspan=""2"">" & vbcrlf
'  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Website Size
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Size:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""website_size"" id=""website_size"" onchange=""setupCustomSize();"">" & vbcrlf
'                                        displayCommunityLinkOptions "WEBSITE_SIZE", lcl_website_size
'  response.write "                    </select>&nbsp;" & vbcrlf
'  response.write "                    <span id=""website_size_customsize_span"">" & vbcrlf
'  response.write "                    <input type=""text"" name=""website_size_customsize"" id=""website_size_customsize"" value=""" & lcl_website_size_customsize & """ size=""5"" maxlength=""10"" onchange=""clearMsg('website_size_customsize');"" />&nbsp;" & vbcrlf
'  response.write "                    <font style=""font-size:10px; color:#800000"">(All sizes in pixels)</font></span>" & vbcrlf
'  response.write "                </td>" & vbcrlf

 'Is Egov Home Page
'  if lcl_isEgovHomePage then
'     lcl_checked_egovhome = " checked=""checked"""
'  else
'     lcl_checked_egovhome = ""
'  end if

'  response.write "                <td align=""right"">" & vbcrlf
'  response.write "                    Make CommunityLink E-Gov Home Page:" & vbcrlf
'  response.write "                    <input type=""checkbox"" name=""isEgovHomePage"" id=""isEgovHomePage"" value=""on""" & lcl_checked_egovhome & " />" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Website Alignment
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Alignment:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <select name=""website_alignment"" id=""website_alignment"">" & vbcrlf
'                                        displayCommunityLinkOptions "WEBSITE_ALIGN", lcl_website_alignment
'  response.write "                    </select>" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Website Background Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Background Color:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'                                      setupColorSelection "website_bgcolor", lcl_website_bgcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('website_bgcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

'  response.write "            <tr><td colspan=""3"">&nbsp;</td></tr>" & vbcrlf

 'Show Logo
'  if lcl_showlogo then
'     lcl_checked_logo = " checked=""checked"""
'  else
'     lcl_checked_logo = ""
'  end if

'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Logo:</td>" & vbcrlf
'  response.write "                <td colspan=""2""><input type=""checkbox"" name=""showlogo"" id=""showlogo"" value=""on""" & lcl_checked_logo & " /></td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Logo Alignment
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Alignment:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <select name=""logo_alignment"" id=""logo_alignment"">" & vbcrlf
'                                        displayCommunityLinkOptions "WEBSITE_LOGO_ALIGN", lcl_logo_alignment
'  response.write "                    </select>" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Logo Filename
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Logo:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <input type=""input"" name=""logo_filename"" id=""logo_filename"" value=""" & lcl_logo_filename & """ size=""50"" maxlength=""500"" onchange=""clearMsg('findImageButton');"" />&nbsp;" & vbcrlf
'  response.write "                    <input type=""button"" name=""findImageButton"" id=""findImageButton"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButton');doPicker('communitylink_maint.logo_filename','Y','','','','N','Y');"" />" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Logo Filename - Background
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Background Logo:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <input type=""input"" name=""logo_filenamebg"" id=""logo_filenamebg"" value=""" & lcl_logo_filenamebg & """ size=""50"" maxlength=""500"" onchange=""clearMsg('findImageButtonbg');"" />&nbsp;" & vbcrlf
'  response.write "                    <input type=""button"" name=""findImageButtonbg"" id=""findImageButtonbg"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButtonbg');doPicker('communitylink_maint.logo_filenamebg','Y','','','','N','Y');"" />" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf
'  response.write "          </table>" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'  response.write "</table>" & vbcrlf

'  if lcl_scripts <> "" then
'     response.write "<script language=""javascript"">" & vbcrlf
'     response.write lcl_scripts & vbcrlf
'     response.write "</script>" & vbcrlf

'     lcl_scripts = ""
'  end if

'  response.write "</div>" & vbcrlf
 'END: Website Logo -----------------------------------------------------------

 'BEGIN: Top Bar Options ------------------------------------------------------
'  response.write "<div id=""tab2"">" & vbcrlf
'  response.write "<table border=""0"" bordercolor=""#00000ff"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""800"" style=""margin-top:10px;"">" & vbcrlf
'  response.write "  <tr>" & vbcrlf
'  response.write "      <td valign=""top"">" & vbcrlf
'  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Top Bar
'  if lcl_showtopbar then
'     lcl_checked_topbar = " checked=""checked"""
'  else
'     lcl_checked_topbar = ""
'  end if

'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Top Bar:</td>" & vbcrlf
'  response.write "                <td><input type=""checkbox"" name=""showtopbar"" id=""showtopbar"" value=""on""" & lcl_checked_topbar & " /></td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Background Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Background Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "topbar_bgcolor", lcl_topbar_bgcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('topbar_bgcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Font Type
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Type:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""topbar_fonttype"" id=""topbar_fonttype"">" & vbcrlf
'                                        displayCommunityLinkOptions "TOPBAR_FONTTYPE", lcl_topbar_fonttype
'  response.write "                    </select>" & vbcrf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Font Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "topbar_fontcolor", lcl_topbar_fontcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('topbar_fontcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Font Color - Hover
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Color<br />(mouseover):</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "topbar_fontcolorhover", lcl_topbar_fontcolorhover, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('topbar_fontcolorhover');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf
'  response.write "          </table>" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'  response.write "</table>" & vbcrlf

'  if lcl_scripts <> "" then
'     response.write "<script language=""javascript"">" & vbcrlf
'     response.write lcl_scripts & vbcrlf
'     response.write "</script>" & vbcrlf

'     lcl_scripts = ""
'  end if

'  response.write "</div>" & vbcrlf
 'END: Top Bar Options --------------------------------------------------------

 'BEGIN: Page Header Options --------------------------------------------------
'  response.write "<div id=""tab3"">" & vbcrlf
'  response.write "<table border=""0"" bordercolor=""#00000ff"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""800"" style=""margin-top:10px;"">" & vbcrlf
'  response.write "  <tr>" & vbcrlf
'  response.write "      <td>" & vbcrlf
'  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Page Header
'  if lcl_showpageheader then
'     lcl_checked_pageheader = " checked=""checked"""
'  else
'     lcl_checked_pageheader = ""
'  end if

'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Page Header:</td>" & vbcrlf
'  response.write "                <td><input type=""checkbox"" name=""showpageheader"" id=""showpageheader"" value=""on""" & lcl_checked_pageheader & " /></td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Background Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Background Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "pageheader_bgcolor", lcl_pageheader_bgcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('pageheader_bgcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Page Header Alignment
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Alignment:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""pageheader_alignment"" id=""pageheader_alignment"">" & vbcrlf
'                                        displayCommunityLinkOptions "PAGEHEADER_ALIGN", lcl_pageheader_alignment
'  response.write "                    </select>" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Page Font Size
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Size:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <input type=""text"" name=""pageheader_fontsize"" id=""pageheader_fontsize"" value=""" & lcl_pageheader_fontsize & """ size=""3"" maxlength=""3"" /> <em style=""font-size:10px; color:#ff0000;"">(in pixels)</em>" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Page Header Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "pageheader_fontcolor", lcl_pageheader_fontcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('pageheader_fontcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Page Header - Font Type
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Type:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""pageheader_fonttype"" id=""pageheader_fonttype"">" & vbcrlf
'                                        displayCommunityLinkOptions "PAGEHEADER_FONTTYPE", lcl_pageheader_fonttype
'  response.write "                    </select>" & vbcrf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

'  response.write "            <tr><td colspan=""3"">&nbsp;</td></tr>" & vbcrlf

 'Show RSS
'  if lcl_showRSS then
'     lcl_checked_showRSS = " checked=""checked"""
'  else
'     lcl_checked_showRSS = ""
'  end if

'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show RSS (icon):</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <input type=""checkbox"" name=""showRSS"" id=""showRSS"" value=""on""" & lcl_checked_showRSS & " />" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'URL - Twitter
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - Twitter:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <input type=""input"" name=""url_twitter"" id=""url_twitter"" value=""" & lcl_url_twitter & """ size=""80"" maxlength=""500"" />" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'URL - Facebook
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - Facebook:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <input type=""input"" name=""url_facebook"" id=""url_facebook"" value=""" & lcl_url_facebook & """ size=""80"" maxlength=""500"" />" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'URL - MySpace
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - MySpace:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <input type=""input"" name=""url_myspace"" id=""url_myspace"" value=""" & lcl_url_myspace & """ size=""80"" maxlength=""500"" />" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'URL - Blogger
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>URL - Blogger:</td>" & vbcrlf
'  response.write "                <td colspan=""2"">" & vbcrlf
'  response.write "                    <input type=""input"" name=""url_blogger"" id=""url_blogger"" value=""" & lcl_url_blogger & """ size=""80"" maxlength=""500"" />&nbsp;" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

'  response.write "          </table>" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'  response.write "</table>" & vbcrlf

'  if lcl_scripts <> "" then
'     response.write "<script language=""javascript"">" & vbcrlf
'     response.write lcl_scripts & vbcrlf
'     response.write "</script>" & vbcrlf

'     lcl_scripts = ""
'  end if

'  response.write "</div>" & vbcrlf
 'END: Page Header Options ----------------------------------------------------

 'BEGIN: Side Menu Bar Options ------------------------------------------------
'  response.write "<div id=""tab4"">" & vbcrlf
'  response.write "<table border=""0"" bordercolor=""#00000ff"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""800"" style=""margin-top:10px;"">" & vbcrlf
'  response.write "  <tr>" & vbcrlf
'  response.write "      <td valign=""top"">" & vbcrlf
'  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Side Menubar
'  if lcl_showsidemenubar then
'     lcl_checked_sidemenubar = " checked=""checked"""
'  else
'     lcl_checked_sidemenubar = ""
'  end if

'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Side Menu Bar:</td>" & vbcrlf
'  response.write "                <td><input type=""checkbox"" name=""showsidemenubar"" id=""showsidemenubar"" value=""on""" & lcl_checked_sidemenubar & " /></td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Side Menubar Alignment
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Alignment:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""sidemenubar_alignment"" id=""sidemenubar_alignment"">" & vbcrlf
'                                        displayCommunityLinkOptions "SIDEMENUBAR_ALIGN", lcl_sidemenubar_alignment
'  response.write "                    </select>" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Side Menubar Option Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Option Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "sidemenuoption_bgcolor", lcl_sidemenuoption_bgcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_bgcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Side Menubar Option Color - Hover
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Option Color<br />(mouseover):</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "sidemenuoption_bgcolorhover", lcl_sidemenuoption_bgcolorhover, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_bgcolorhover');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

'  response.write "            <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf

 'Side Menubar Option Alignment
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Text Alignment:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""sidemenuoption_alignment"" id=""sidemenuoption_alignment"">" & vbcrlf
'                                        displayCommunityLinkOptions "SIDEMENUOPT_TEXTALIGN", lcl_sidemenuoption_alignment
'  response.write "                    </select>" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Side Menubar Option - Font Type
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Option Font Type:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""sidemenuoption_fonttype"" id=""sidemenuoption_fonttype"">" & vbcrlf
'                                        displayCommunityLinkOptions "SIDEMENUOPT_FONTTYPE", lcl_sidemenuoption_fonttype
'  response.write "                    </select>" & vbcrf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Side Menubar Option - Font Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Option Font Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "sidemenuoption_fontcolor", lcl_sidemenuoption_fontcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_fontcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Side Menubar Option - Font Color - Hover
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Option Font Color<br />(mouseover):</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "sidemenuoption_fontcolorhover", lcl_sidemenuoption_fontcolorhover, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_fontcolorhover');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf
'  response.write "          </table>" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'  response.write "</table>" & vbcrlf

'  if lcl_scripts <> "" then
'     response.write "<script language=""javascript"">" & vbcrlf
'     response.write lcl_scripts & vbcrlf
'     response.write "</script>" & vbcrlf

'     lcl_scripts = ""
'  end if

'  response.write "</div>" & vbcrlf
 'END: Side Menu Bar Options --------------------------------------------------

 'BEGIN: Footer Options -------------------------------------------------------
'  response.write "<div id=""tab5"">" & vbcrlf
'  response.write "<table border=""0"" bordercolor=""#00000ff"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""800"" style=""margin-top:10px;"">" & vbcrlf
'  response.write "  <tr>" & vbcrlf
'  response.write "      <td valign=""top"">" & vbcrlf
'  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Footer
'  if lcl_showfooter then
'     lcl_checked_footer = " checked=""checked"""
'  else
'     lcl_checked_footer = ""
'  end if

'  response.write "            <tr>" & vbcrlf
'  response.write "                <td width=""" & lcl_labelcolumn_width & """>Show Footer:</td>" & vbcrlf
'  response.write "                <td><input type=""checkbox"" name=""showfooter"" id=""showfooter"" value=""on""" & lcl_checked_footer & " /></td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Background Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Background Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "footer_bgcolor", lcl_footer_bgcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('footer_bgcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Font Type
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Type:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'  response.write "                    <select name=""footer_fonttype"" id=""footer_fonttype"">" & vbcrlf
'                                        displayCommunityLinkOptions "FOOTER_FONTTYPE", lcl_footer_fonttype
'  response.write "                    </select>" & vbcrf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Font Color
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Color:</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "footer_fontcolor", lcl_footer_fontcolor, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('footer_fontcolor');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Font Color - Hover
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td>Font Color<br />(mouseover):</td>" & vbcrlf
'  response.write "                <td>" & vbcrlf
'                                      setupColorSelection "footer_fontcolorhover", lcl_footer_fontcolorhover, 1
'                                      lcl_scripts = lcl_scripts & "changePreviewColor('footer_fontcolorhover');" & vbcrlf
'  response.write "                </td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf
'  response.write "          </table>" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'  response.write "</table>" & vbcrlf

'  if lcl_scripts <> "" then
'     response.write "<script language=""javascript"">" & vbcrlf
'     response.write lcl_scripts & vbcrlf
'     response.write "</script>" & vbcrlf

'     lcl_scripts = ""
'  end if

'  response.write "</div>" & vbcrlf
 'END: Footer Options ---------------------------------------------------------

 'BEGIN: CommunityLink Options ------------------------------------------------
'  iBGColor              = "#eeeeee"
'  iRowCount             = 0
'  lcl_features_shown    = ""
'  iTotalCLFeaturesAvail = getCLFeatAvailCount(session("orgid"))

'  response.write "<div id=""tab6"">" & vbcrlf
'  response.write "<table border=""0"" bordercolor=""#00000ff"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""800"" style=""margin-top:10px;"">" & vbcrlf
'  response.write "  <tr>" & vbcrlf
'  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
'  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""3"" class=""communitylink_table"">" & vbcrlf
'  response.write "            <tr>" & vbcrlf
'  response.write "                <th align=""left"">CommunityLink Sections</th>" & vbcrlf
'  response.write "                <th nowrap=""nowrap"">Community Link</th>" & vbcrlf
'  response.write "                <th nowrap=""nowrap"">Savvy/IFRAME</th>" & vbcrlf
'  response.write "            </tr>" & vbcrlf

 'Setup the defaults.
'  lcl_isCommunityLinkOn       = 0
'  lcl_isSavvyOn               = 0
'  lcl_showsectionborder       = 0
'  lcl_sectionbordercolor      = "000000"
'  lcl_sectionheader_bgcolor   = "ffffff"
'  lcl_sectionheader_linecolor = "000000"
'  lcl_sectionheader_fonttype  = getCLOptionDefault("SECTIONHEADER_FONTTYPE")
'  lcl_sectionheader_fontcolor = "000000"
'  lcl_sectionheader_fontsize  = "11"
'  lcl_sectionheader_isbold    = 1
'  lcl_sectionheader_isitalic  = 0
'  lcl_sectiontext_bgcolor     = "ffffff"
'  lcl_sectiontext_fonttype    = getCLOptionDefault("SECTIONTEXT_FONTTYPE")
'  lcl_sectiontext_fontcolor   = "000000"
'  lcl_sectionlinks_alignment  = getCLOptionDefault("SECTIONLINKS_ALIGN")
'  lcl_sectionlinks_fonttype   = getCLOptionDefault("SECTIONLINKS_FONTTYPE")
'  lcl_sectionlinks_fontcolor  = "800000"

'  sSQL = " SELECT f.featureid, "
'  sSQL = sSQL & " isnull(f.CL_portaltype,'') AS portaltype, "
'  sSQL = sSQL & " f.CommunityLinkOn, "
'  sSQL = sSQL & " isnull(cl.featurename, isnull(otf.featurename, f.featurename)) AS featurename, "
'  sSQL = sSQL & " isnull(otf.featurename, f.featurename) AS featurename_original, "
'  sSQL = sSQL & " isnull(cl.portalcolumn, 1) AS portalcolumn, "
'  sSQL = sSQL & " isnull(cl.displayorder, 1) AS displayorder, "
'  sSQL = sSQL & " cl.rss_feedid, "
'  sSQL = sSQL & " isnull(cl.numListItemsShown_CL, f.CL_numListItems) AS numListItemsShown_CL, "
'  sSQL = sSQL & " isnull(cl.numListItemsShown_SAVVY, f.CL_numListItems) AS numListItemsShown_SAVVY, "
'  sSQL = sSQL & " f.CL_numListItems AS numListItemsShown_original, "
'  sSQL = sSQL & " isnull(cl.isCommunityLinkOn,"                  & lcl_isCommunityLinkOn       & ") AS isCommunityLinkOn, "
'  sSQL = sSQL & " isnull(cl.isSavvyOn,"                          & lcl_isSavvyOn               & ") AS isSavvyOn, "
'  sSQL = sSQL & " isnull(cl.showsectionborder_CL,"               & lcl_showsectionborder       & ") AS showsectionborder_CL, "
'  sSQL = sSQL & " isnull(cl.showsectionborder_SAVVY,"            & lcl_showsectionborder       & ") AS showsectionborder_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionbordercolor_CL,'"             & lcl_sectionbordercolor      & "') AS sectionbordercolor_CL, "
'  sSQL = sSQL & " isnull(cl.sectionbordercolor_SAVVY,'"          & lcl_sectionbordercolor      & "') AS sectionbordercolor_SAVVY, "
'  sSQL = sSQL & " cl.sectionbackgroundcolor_CL, "
'  sSQL = sSQL & " cl.sectionbackgroundcolor_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionheader_bgcolor_CL,'"          & lcl_sectionheader_bgcolor   & "') AS sectionheader_bgcolor_CL, "
'  sSQL = sSQL & " isnull(cl.sectionheader_bgcolor_SAVVY,'"       & lcl_sectionheader_bgcolor   & "') AS sectionheader_bgcolor_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionheader_linecolor_CL,'"        & lcl_sectionheader_linecolor & "') AS sectionheader_linecolor_CL, "
'  sSQL = sSQL & " isnull(cl.sectionheader_linecolor_SAVVY,'"     & lcl_sectionheader_linecolor & "') AS sectionheader_linecolor_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionheader_fonttype_CL,'"         & lcl_sectionheader_fonttype  & "') AS sectionheader_fonttype_CL, "
'  sSQL = sSQL & " isnull(cl.sectionheader_fonttype_SAVVY,'"      & lcl_sectionheader_fonttype  & "') AS sectionheader_fonttype_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionheader_fontcolor_CL,'"        & lcl_sectionheader_fontcolor & "') AS sectionheader_fontcolor_CL, "
'  sSQL = sSQL & " isnull(cl.sectionheader_fontcolor_SAVVY,'"     & lcl_sectionheader_fontcolor & "') AS sectionheader_fontcolor_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionheader_fontsize_CL,'"         & lcl_sectionheader_fontsize  & "') AS sectionheader_fontsize_CL, "
'  sSQL = sSQL & " isnull(cl.sectionheader_fontsize_SAVVY,'"      & lcl_sectionheader_fontsize  & "') AS sectionheader_fontsize_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionheader_isbold_CL,'"           & lcl_sectionheader_isbold    & "') AS sectionheader_isbold_CL, "
'  sSQL = sSQL & " isnull(cl.sectionheader_isbold_SAVVY,'"        & lcl_sectionheader_isbold    & "') AS sectionheader_isbold_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionheader_isitalic_CL,'"         & lcl_sectionheader_isitalic  & "') AS sectionheader_isitalic_CL, "
'  sSQL = sSQL & " isnull(cl.sectionheader_isitalic_SAVVY,'"      & lcl_sectionheader_isitalic  & "') AS sectionheader_isitalic_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectiontext_bgcolor_CL,'"            & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolor_CL, "
'  sSQL = sSQL & " isnull(cl.sectiontext_bgcolor_SAVVY,'"         & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolor_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectiontext_bgcolorhover_CL,'"       & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolorhover_CL, "
'  sSQL = sSQL & " isnull(cl.sectiontext_bgcolorhover_SAVVY,'"    & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolorhover_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fonttype_CL,'"           & lcl_sectiontext_fonttype    & "') AS sectiontext_fonttype_CL, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fonttype_SAVVY,'"        & lcl_sectiontext_fonttype    & "') AS sectiontext_fonttype_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fontcolor_CL,'"          & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolor_CL, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fontcolor_SAVVY,'"       & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolor_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fontcolorhover_CL,'"     & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolorhover_CL, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fontcolorhover_SAVVY,'"  & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolorhover_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fontsize_CL,'"           & lcl_sectiontext_fontsize    & "') AS sectiontext_fontsize_CL, "
'  sSQL = sSQL & " isnull(cl.sectiontext_fontsize_SAVVY,'"        & lcl_sectiontext_fontsize    & "') AS sectiontext_fontsize_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_alignment_CL,'"         & lcl_sectionlinks_alignment  & "') AS sectionlinks_alignment_CL, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_alignment_SAVVY,'"      & lcl_sectionlinks_alignment  & "') AS sectionlinks_alignment_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_fonttype_CL,'"          & lcl_sectionlinks_fonttype   & "') AS sectionlinks_fonttype_CL, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_fonttype_SAVVY,'"       & lcl_sectionlinks_fonttype   & "') AS sectionlinks_fonttype_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolor_CL,'"         & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolor_CL, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolor_SAVVY,'"      & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolor_SAVVY, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolorhover_CL,'"    & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolorhover_CL, "
'  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolorhover_SAVVY,'" & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolorhover_SAVVY, "
'  sSQL = sSQL & " query_filter "
'  sSQL = sSQL & " FROM egov_communitylink_displayorgfeatures cl "
'  sSQL = sSQL &      " RIGHT OUTER JOIN egov_organizations_to_features otf "
'  sSQL = sSQL &      " INNER JOIN egov_organization_features f "
'  sSQL = sSQL &      " ON otf.featureid = f.featureid "
'  sSQL = sSQL &      " ON f.featureid = cl.featureid "
'  sSQL = sSQL &      " AND cl.orgid = otf.orgid "
'  sSQL = sSQL & " WHERE f.haspublicview = 1 "
'  sSQL = sSQL & " AND f.CommunityLinkOn = 1 "
'  sSQL = sSQL & " AND otf.orgid = " & session("orgid")
'  sSQL = sSQL & " ORDER BY cl.isCommunityLinkOn DESC, isnull(cl.portalcolumn, 0), isnull(cl.displayorder, 0), "
'  sSQL = sSQL &          " isnull(cl.featurename, isnull(otf.featurename, f.featurename)), cl.isSavvyOn DESC "

'  set oCLFeatures = Server.CreateObject("ADODB.Recordset")
'  oCLFeatures.Open sSQL, Application("DSN"), 3, 1

'  if not oCLFeatures.eof then
'     do while not oCLFeatures.eof
'        iBGColor  = changeBGColor(iBGColor,"#eeeeee","#ffffff")
'        iRowCount = iRowCount + 1

'        response.write "            <tr valign=""top"" bgcolor=""" & iBGColor & """>" & vbcrlf

'        displayCommunityLinkFeatureOptions "CL", _
'                                           iRowCount, _
'                                           iTotalCLFeaturesAvail, _
'                                           iBGColor, _
'                                           lcl_scripts, _
'                                           oCLFeatures("featureid"), _
'                                           oCLFeatures("featurename"), _
'                                           oCLFeatures("featurename_original"), _
'                                           oCLFeatures("portalcolumn"), _
'                                           oCLFeatures("displayorder"), _
'                                           oCLFeatures("rss_feedid"), _
'                                           oCLFeatures("numListItemsShown_CL"), _
'                                           oCLFeatures("numListItemsShown_original"), _
'                                           oCLFeatures("isCommunityLinkOn"), _
'                                           oCLFeatures("showsectionborder_CL"), _
'                                           oCLFeatures("sectionbordercolor_CL"), _
'                                           oCLFeatures("sectionbackgroundcolor_CL"), _
'                                           oCLFeatures("sectionheader_bgcolor_CL"), _
'                                           oCLFeatures("sectionheader_linecolor_CL"), _
'                                           oCLFeatures("sectionheader_fonttype_CL"), _
'                                           oCLFeatures("sectionheader_fontcolor_CL"), _
'                                           oCLFeatures("sectionheader_fontsize_CL"), _
'                                           oCLFeatures("sectionheader_isbold_CL"), _
'                                           oCLFeatures("sectionheader_isitalic_CL"), _
'                                           oCLFeatures("sectiontext_bgcolor_CL"), _
'                                           oCLFeatures("sectiontext_bgcolorhover_CL"), _
'                                           oCLFeatures("sectiontext_fonttype_CL"), _
'                                           oCLFeatures("sectiontext_fontcolor_CL"), _
'                                           oCLFeatures("sectiontext_fontcolorhover_CL"), _
'                                           oCLFeatures("sectiontext_fontsize_CL"), _
'                                           oCLFeatures("sectionlinks_alignment_CL"), _
'                                           oCLFeatures("sectionlinks_fonttype_CL"), _
'                                           oCLFeatures("sectionlinks_fontcolor_CL"), _
'                                           oCLFeatures("sectionlinks_fontcolorhover_CL"), _
'                                           oCLFeatures("query_filter")

'        displayCommunityLinkFeatureOptions "SAVVY", _
'                                           iRowCount, _
'                                           iTotalCLFeaturesAvail, _
'                                           iBGColor, _
'                                           lcl_scripts, _
'                                           oCLFeatures("featureid"), _
'                                           oCLFeatures("featurename"), _
'                                           oCLFeatures("featurename_original"), _
'                                           oCLFeatures("portalcolumn"), _
'                                           oCLFeatures("displayorder"), _
'                                           oCLFeatures("rss_feedid"), _
'                                           oCLFeatures("numListItemsShown_SAVVY"), _
'                                           oCLFeatures("numListItemsShown_original"), _
'                                           oCLFeatures("isSavvyOn"), _
'                                           oCLFeatures("showsectionborder_SAVVY"), _
'                                           oCLFeatures("sectionbordercolor_SAVVY"), _
'                                           oCLFeatures("sectionbackgroundcolor_SAVVY"), _
'                                           oCLFeatures("sectionheader_bgcolor_SAVVY"), _
'                                           oCLFeatures("sectionheader_linecolor_SAVVY"), _
'                                           oCLFeatures("sectionheader_fonttype_SAVVY"), _
'                                           oCLFeatures("sectionheader_fontcolor_SAVVY"), _
'                                           oCLFeatures("sectionheader_fontsize_SAVVY"), _
'                                           oCLFeatures("sectionheader_isbold_SAVVY"), _
'                                           oCLFeatures("sectionheader_isitalic_SAVVY"), _
'                                           oCLFeatures("sectiontext_bgcolor_SAVVY"), _
'                                           oCLFeatures("sectiontext_bgcolorhover_SAVVY"), _
'                                           oCLFeatures("sectiontext_fonttype_SAVVY"), _
'                                           oCLFeatures("sectiontext_fontcolor_SAVVY"), _
'                                           oCLFeatures("sectiontext_fontcolorhover_SAVVY"), _
'                                           oCLFeatures("sectiontext_fontsize_SAVVY"), _
'                                           oCLFeatures("sectionlinks_alignment_SAVVY"), _
'                                           oCLFeatures("sectionlinks_fonttype_SAVVY"), _
'                                           oCLFeatures("sectionlinks_fontcolor_SAVVY"), _
'                                           oCLFeatures("sectionlinks_fontcolorhover_SAVVY"), _
'                                           oCLFeatures("query_filter")

'        response.write "            </tr>" & vbcrlf

'        oCLFeatures.movenext
'     loop
'  end if

'  oCLFeatures.close
'  set oCLFeatures = nothing

 'Total Rows
'  response.write "            <tr>" & vbcrlf
'  response.write "                <td><input type=""hidden"" name=""totalCLRows"" id=""totalCLRows"" value=""" & iRowCount & """ size=""3"" maxlength=""10"" /></td>" & vbcrlf
'  response.write "            </tr>" & vbcrlf
'  response.write "          </table>" & vbcrlf
'  response.write "      </td>" & vbcrlf
'  response.write "  </tr>" & vbcrlf
'  response.write "</table>" & vbcrlf
'  response.write "  </form>" & vbcrlf
'  response.write "</div>" & vbcrlf
 'END: CommunityLink Options --------------------------------------------------

 'BEGIN: Close TABS -----------------------------------------------------------
  'response.write "</div>" & vbcrlf
 'EMD: Close TABS -------------------------------------------------------------

  displayButtons "MAINT"

  'response.write "  </div>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"--> 
<%
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf

     lcl_scripts = ""
  end if

  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub displayButtons(iType)

  if iTYPE = "PREVIEW" then
     response.write "<input type=""button"" name=""previewButton"" id=""previewButton"" value=""Refresh Preview"" class=""button"" onclick=""location.href='communitylink_maint.asp';"" />" & vbcrlf
  else
     response.write "<input type=""button"" name=""saveChanges"" id=""saveChanges"" value=""Save Changes"" class=""button"" onclick=""return validateFields()"" />" & vbcrlf
  end if

end sub



'------------------------------------------------------------------------------
sub displayCommunityLinkFeatureOptions(p_rowType, _
                                       p_rowcount, _
                                       p_totalrows, _
                                       p_bgcolor, _
                                       p_scripts, _
                                       p_featureid, _
                                       p_featurename, _
                                       p_featurename_original, _
                                       p_portalcolumn, _
                                       p_displayorder, _
                                       p_rss_feedid, _
                                       p_numListItemsShown, _
                                       p_numListItemsShown_original, _
                                       p_showSection, _
                                       p_showSectionBorder, _
                                       p_sectionBorderColor, _
                                       p_sectionBackgroundColor, _
                                       p_sectionheader_bgcolor, _
                                       p_sectionheader_linecolor, _
                                       p_sectionheader_fonttype, _
                                       p_sectionheader_fontcolor, _
                                       p_sectionheader_fontsize, _
                                       p_sectionheader_isbold, _
                                       p_sectionheader_isitalic, _
                                       p_sectiontext_bgcolor, _
                                       p_sectiontext_bgcolorhover, _
                                       p_sectiontext_fonttype, _
                                       p_sectiontext_fontcolor, _
                                       p_sectiontext_fontcolorhover, _
                                       p_sectiontext_fontsize, _
                                       p_sectionlinks_align, _
                                       p_sectionlinks_fonttype, _
                                       p_sectionlinks_fontcolor, _
                                       p_sectionlinks_fontcolorhover, _
                                       p_viewall_urltype, _
                                       p_viewall_url, _
                                       p_viewall_url_wintype, _
                                       p_query_filter)

  if p_rowcount > 1 then
     lcl_row_border = "border-top:1pt solid #000000;"
  else
     lcl_row_border = ""
  end if

  if p_rowType <> "" then
     lcl_rowType = UCASE(p_rowType)
  else
     lcl_rowType = "CL"
  end if

 'Setup the script that will determine if this row is enabled/disable when the screen is opened.
  'lcl_scripts = p_scripts & "enableDisableOptions('" & p_rowType & "'," & p_rowcount & ");" & vbcrlf
  lcl_scripts = p_scripts

 'Determine if this feature is "turned-on" to be displayed on this org's Community Link screen.
  if p_showSection then
     lcl_checked_showSection = " checked=""checked"""
  else
     lcl_checked_showSection = ""
  end if

 'Determine if the "section border" is turned-on
  if p_showSectionBorder then
     lcl_checked_showSectionBorder = " checked=""checked"""
  else
     lcl_checked_showSectionBorder = ""
  end if

 'Determine if the "section header" is bold/italic
  if p_sectionheader_isbold then
     lcl_checked_sectionHeaderIsBold = " checked=""checked"""
  else
     lcl_checked_sectionHeaderIsBold = ""
  end if

  if p_sectionheader_isitalic then
     lcl_checked_sectionHeaderIsItalic = " checked=""checked"""
  else
     lcl_checked_sectionHeaderIsItalic = ""
  end if

  if lcl_rowType = "SAVVY" then
     lcl_td_style         = "border-left:1pt solid #000000;"
     lcl_showSectionLabel = "Style"
  else
     lcl_td_style         = ""
     lcl_showSectionLabel = "Display"
  end if

  if lcl_rowType = "CL" then
     response.write "                  <td style=""padding-left:10px;" & lcl_row_border & """>" & vbcrlf
     response.write "                      <input type=""hidden"" name=""featureid_"                  & p_rowcount & """ id=""featureid_"                  & p_rowcount & """ value=""" & p_featureid                  & """ size=""5"" maxlength=""10"" />" & vbcrlf
     response.write "                      <input type=""hidden"" name=""featurename_original_"       & p_rowcount & """ id=""featurename_original_"       & p_rowcount & """ value=""" & p_featurename_original       & """ size=""5"" maxlength=""255"" />" & vbcrlf
     response.write "                      <input type=""hidden"" name=""numListItemsShown_original_" & p_rowcount & """ id=""numListItemsShown_original_" & p_rowcount & """ value=""" & p_numListItemsShown_original & """ size=""3"" maxlength=""10"" />" & vbcrlf
     response.write "                      <input type=""hidden"" name=""editCL_"                     & p_rowcount & """ id=""editCL_"                     & p_rowcount & """ value=""N"" size=""3"" maxlength=""3"" />" & vbcrlf
     response.write "                      <input type=""hidden"" name=""editSAVVY_"                  & p_rowcount & """ id=""editSAVVY_"                  & p_rowcount & """ value=""N"" size=""3"" maxlength=""3"" />" & vbcrlf

    'Feature Name
     response.write "                      <input type=""text"" name=""featurename_" & p_rowcount & """ id=""featurename_" & p_rowcount & """ value=""" & p_featurename & """ size=""40"" maxlength=""255"" /><br />" & vbcrlf
     response.write "                      <table id=""sectionTable_" & p_rowcount & """ class=""sectionTable"" border=""0"" cellspacing=""0"" cellpadding=""2"" align=""right"" style=""background-color:" & p_bgcolor & "; margin-left:10px; margin-top:5px;"">" & vbcrlf

    'Portal (Display) Column
     response.write "                        <tr id=""styleProperties_Portal_" & p_rowcount & """>" & vbcrlf
     response.write "                            <td width=""100"">Display Column:</td>" & vbcrlf
     response.write "                            <td>" & vbcrlf
     response.write "                                <select name=""portalcolumn_" & p_rowcount & """ id=""portalcolumn_" & p_rowcount & """>" & vbcrlf
                                                       displayCommunityLinkOptions "PORTALCOLUMNS", p_portalcolumn
     response.write "                                </select>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf

    'Display (Row) Order
     response.write "                        <tr id=""styleProperties_DisplayOrder_" & p_rowcount & """>" & vbcrlf
     response.write "                            <td>Display Order:</td>" & vbcrlf
     response.write "                            <td><input type=""text"" name=""displayorder_" & p_rowcount & """ id=""displayorder_" & p_rowcount & """ value=""" & p_displayorder & """ size=""3"" maxlength=""5"" onchange=""clearMsg('displayorder_" & p_rowcount & "');"" /></td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf

    'RSS Feed - FeatureID
     response.write "                        <tr id=""styleProperties_RSSFeedFeature_" & p_rowcount & """ valign=""top"">" & vbcrlf
     response.write "                            <td>RSS Feed:<br />(feature)</td>" & vbcrlf
     response.write "                            <td>" & vbcrlf
     response.write "                                <select name=""rss_feedid_" & p_rowcount & """ id=""rss_feedid_" & p_rowcount & """>" & vbcrlf
                                                       displayRSSFeedOptions p_rss_feedid
     response.write "                                </select>" & vbcrlf
     response.write "                            </td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf

    'Reset Button
     response.write "                        <tr id=""styleProperties_ResetButton_" & p_rowcount & """>" & vbcrlf
     response.write "                            <td>Reset to Defaults:&nbsp;" & vbcrlf
     response.write "                            <td><input type=""button"" name=""resetButton_" & p_rowcount & """ id=""resetButton_" & p_rowcount & """ value=""Reset"" class=""button"" onclick=""resetFields('" & p_rowcount & "');"" /></td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf
     response.write "                      </table>" & vbcrlf

    'Filter
     'response.write "                        <tr id=""styleProperties_filter_" & p_rowcount & """>" & vbcrlf
     'response.write "                            <td colspan=""2"">" & vbcrlf
     response.write "                                <div id=""queryfilter_div_" & p_rowcount & """ class=""queryfilter_div"">" & vbcrlf
     response.write "                                  Filter Results: (SQL)<br />" & vbcrlf
     response.write "                                  <textarea name=""query_filter_" & p_rowcount & """ id=""query_filter_" & p_rowcount & """ rows=""10"" cols=""40"">" & p_query_filter & "</textarea><br />" & vbcrlf
     response.write "                                  <span class=""redText"">NOTE: This filter will be used within the WHERE clause of the SQL query.  This filter <strong>MUST</strong> begin with (<strong>AND</strong>).</span>" & vbcrlf
     response.write "                                </div>" & vbcrlf
     'response.write "                            </td>" & vbcrlf
     'response.write "                        </tr>" & vbcrlf

     response.write "                  </td>" & vbcrlf
  end if

  response.write "                  <td style=""color:#800000;padding-left:10px;" & lcl_td_style & lcl_row_border & """ nowrap=""nowrap"">" & vbcrlf

 'Show Section
  'response.write "                      <div>" & lcl_showSectionLabel & ":&nbsp;" & vbcrlf
  'response.write "                        <input type=""checkbox"" name=""showSection_" & lcl_rowType & "_" & p_rowcount & """ id=""showSection_" & lcl_rowType & "_" & p_rowcount & """ value=""on"" onclick=""enableDisableOptions('" & p_rowType & "','" & p_rowcount & "');""" & lcl_checked_showSection & " />" & vbcrlf
  response.write "                       <div>" & vbcrlf
  'response.write "                        <input type=""checkbox"" name=""showSection_" & lcl_rowType & "_" & p_rowcount & """ id=""showSection_" & lcl_rowType & "_" & p_rowcount & """ value=""on""" & lcl_checked_showSection & " />" & vbcrlf
  response.write "                        <input type=""button"" name=""editButton_" & lcl_rowType & "_" & p_rowcount & """ id=""editButton_" & lcl_rowType & "_" & p_rowcount & """ value=""Edit"" class=""button"" onclick=""showHideOptions('S','" & p_rowType & "','" & p_rowcount & "');"" />" & vbcrlf
  response.write "                        <input type=""button"" name=""closeButton_" & lcl_rowType & "_" & p_rowcount & """ id=""closeButton_" & lcl_rowType & "_" & p_rowcount & """ value=""Finished Editing"" class=""closeButton"" onclick=""showHideOptions('H', '" & p_rowType & "','" & p_rowcount & "');"" />" & vbcrlf
  response.write "                        <img src=""../images/help.jpg"" name=""helpFeature_edit_" & lcl_rowType & "_" & p_rowcount & """ id=""helpFeature_edit_" & lcl_rowType & "_" & p_rowcount & """ class=""helpOption"" alt=""Click for more info"" onclick=""showHideHelp('" & p_rowType & "','" & p_rowcount & "')"" /><br />" & vbcrlf
  response.write "                        <div name=""helpFeature_edit_" & lcl_rowType & "_" & p_rowcount & "_text"" id=""helpFeature_edit_" & lcl_rowType & "_" & p_rowcount & "_text"" class=""helpOptionText"">" & vbcrlf
  response.write "                          <p><strong>E-GOV TIP:</strong><br />Any/All changes are NOT saved until the ""Save Changes"" button is clicked.<br />The ""Finished Editing"" button simply hides the options.</p>" & vbcrlf
  response.write "                        </div>" & vbcrlf
  response.write "                      </div>" & vbcrlf
  response.write "                      <fieldset class=""fieldset"" id=""sectionFieldset_" & lcl_rowType & "_" & p_rowcount & """>" & vbcrlf

  'response.write "                      <table id=""styleProperties_" & lcl_rowType & "_" & p_rowcount & """ border=""0"" cellspacing=""0"" cellpadding=""2"" style=""background-color:" & p_bgcolor & "; margin-top:2px; margin-left:4px;"">" & vbcrlf
  response.write "                      <table id=""styleProperties_" & lcl_rowType & "_" & p_rowcount & """ border=""0"" cellspacing=""0"" cellpadding=""2"" style=""margin: 2px 0px 0px 4px;"">" & vbcrlf

 'Display
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & lcl_showSectionLabel & ":</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap""><input type=""checkbox"" name=""showSection_" & lcl_rowType & "_" & p_rowcount & """ id=""showSection_" & lcl_rowType & "_" & p_rowcount & """ value=""on""" & lcl_checked_showSection & " /></td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Number of List Items Shown
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap""># List Items:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap""><input type=""text"" name=""numListItemsShown_" & lcl_rowType & "_" & p_rowcount & """ id=""numListItemsShown_" & lcl_rowType & "_" & p_rowcount & """ value=""" & p_numListItemsShown & """ size=""3"" maxlength=""10"" onchange=""clearMsg('numListItemsShown_" & lcl_rowType & "_" & p_rowcount & "');"" /></td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Display Section Border
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Show Border:</td>" & vbcrlf
  'response.write "                            <td nowrap=""nowrap""><input type=""checkbox"" name=""showsectionborder_" & lcl_rowType & "_" & p_rowcount & """ id=""showsectionborder_" & lcl_rowType & "_" & p_rowcount & """ value=""on"" onclick=""enableDisableOptions('" & p_rowType & "','" & p_rowcount & "');""" & lcl_checked_showSectionBorder & " /></td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap""><input type=""checkbox"" name=""showsectionborder_" & lcl_rowType & "_" & p_rowcount & """ id=""showsectionborder_" & lcl_rowType & "_" & p_rowcount & """ value=""on""" & lcl_checked_showSectionBorder & " /></td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Border Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">Border Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">" & vbcrlf

                                                  setupColorSelection "sectionbordercolor_" & lcl_rowType & "_" & p_rowcount, p_sectionBorderColor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectionbordercolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf

  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Background Color (BODY override for iFrames ONLY)
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">" & vbcrlf
  response.write "                                Background Color:<br />" & vbcrlf
  response.write "                                (iFrame ONLY)" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">" & vbcrlf

                                                  setupColorSelection "sectionbackgroundcolor_" & lcl_rowType & "_" & p_rowcount, p_sectionBackgroundColor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectionbackgroundcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf

  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf


 'BEGIN: Header Options -------------------------------------------------------
 'Section Header - Background Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Header BG Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionheader_bgcolor_" & lcl_rowType & "_" & p_rowcount, p_sectionheader_bgcolor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectionheader_bgcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Header - Line Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Line BG Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionheader_linecolor_" & lcl_rowType & "_" & p_rowcount, p_sectionheader_linecolor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectionheader_linecolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Header - Font Type
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Header Font Type:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
  response.write "                                <select name=""sectionheader_fonttype_" & lcl_rowType & "_" & p_rowcount & """ id=""sectionheader_fonttype_" & lcl_rowType & "_" & p_rowcount & """>" & vbcrlf
                                                    displayCommunityLinkOptions "SECTIONHEADER_FONTTYPE", p_sectionheader_fonttype
  response.write "                                </select>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Header - Font Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Header Font Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionheader_fontcolor_" & lcl_rowType & "_" & p_rowcount, p_sectionheader_fontcolor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectionheader_fontcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Header - Font Size
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">Header Font Size:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">" & vbcrlf
  response.write "                                <input type=""text"" name=""sectionheader_fontsize_" & lcl_rowType & "_" & p_rowcount & """ id=""sectionheader_fontsize_" & lcl_rowType & "_" & p_rowcount & """ value=""" & p_sectionheader_fontsize & """ size=""3"" maxlength=""3"" onchange=""clearMsg('sectionheader_fontsize_" & lcl_rowType & "_" & p_rowcount & "');"" />&nbsp;&nbsp;" & vbcrlf
  response.write "                                <input type=""checkbox"" name=""sectionheader_isbold_" & lcl_rowType & "_" & p_rowcount & """ id=""sectionheader_isbold_" & lcl_rowType & "_" & p_rowcount & """ value=""on""" & lcl_checked_sectionHeaderIsBold & " /> Bold&nbsp;" & vbcrlf
  response.write "                                <input type=""checkbox"" name=""sectionheader_isitalic_" & lcl_rowType & "_" & p_rowcount & """ id=""sectionheader_isitalic_" & lcl_rowType & "_" & p_rowcount & """ value=""on""" & lcl_checked_sectionHeaderIsItalic & " /> Italic" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf


 'END: Header Options ---------------------------------------------------------

 'BEGIN: Section Text Options -------------------------------------------------
 'Section Text - Background Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Text BG Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectiontext_bgcolor_" & lcl_rowType & "_" & p_rowcount, p_sectiontext_bgcolor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectiontext_bgcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Text - Background Color Hover
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Text BG Color<br />(mouseover):</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectiontext_bgcolorhover_" & lcl_rowType & "_" & p_rowcount, p_sectiontext_bgcolorhover, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectiontext_bgcolorhover_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Text - Font Type
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Text Font Type:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
  response.write "                                <select name=""sectiontext_fonttype_" & lcl_rowType & "_" & p_rowcount & """ id=""sectiontext_fonttype_" & lcl_rowType & "_" & p_rowcount & """>" & vbcrlf
                                                    displayCommunityLinkOptions "SECTIONTEXT_FONTTYPE", p_sectiontext_fonttype
  response.write "                                </select>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Text - Font Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Text Font Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectiontext_fontcolor_" & lcl_rowType & "_" & p_rowcount, p_sectiontext_fontcolor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectiontext_fontcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Text - Font Color Hover
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Text Font Color:<br />(mouseover)</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectiontext_fontcolorhover_" & lcl_rowType & "_" & p_rowcount, p_sectiontext_fontcolorhover, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectiontext_fontcolorhover_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Header - Font Size
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">Section Text - Font Size:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">" & vbcrlf
  response.write "                                <input type=""text"" name=""sectiontext_fontsize_" & lcl_rowType & "_" & p_rowcount & """ id=""sectiontext_fontsize_" & lcl_rowType & "_" & p_rowcount & """ value=""" & p_sectiontext_fontsize & """ size=""3"" maxlength=""3"" onchange=""clearMsg('sectiontext_fontsize_" & lcl_rowType & "_" & p_rowcount & "');"" />&nbsp;&nbsp;" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
 'END: Section Text Options ---------------------------------------------------

 'BEGIN: Link Row Options -----------------------------------------------------
 'Link Row - Alignment
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td>Link Row - Alignment:</td>" & vbcrlf
  response.write "                            <td colspan=""2"">" & vbcrlf
  response.write "                                <select name=""sectionlinks_alignment_" & lcl_rowType & "_" & p_rowcount & """ id=""sectionlinks_alignment_" & lcl_rowType & "_" & p_rowcount & """>" & vbcrlf
                                                    displayCommunityLinkOptions "SECTIONLINKS_ALIGN", p_sectionlinks_align
  response.write "                                </select>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Link Row - Font Type
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Link Row - Font Type:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
  response.write "                                <select name=""sectionlinks_fonttype_" & lcl_rowType & "_" & p_rowcount & """ id=""sectionlinks_fonttype_" & lcl_rowType & "_" & p_rowcount & """>" & vbcrlf
                                                    displayCommunityLinkOptions "SECTIONLINKS_FONTTYPE", p_sectiontext_fonttype
  response.write "                                </select>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Link Row - Font Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Link Row - Font Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionlinks_fontcolor_" & lcl_rowType & "_" & p_rowcount, p_sectionlinks_fontcolor, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectionlinks_fontcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Link Row - Font Color Hover
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Link Row - Font Color:<br />(mouseover)</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionlinks_fontcolorhover_" & lcl_rowType & "_" & p_rowcount, p_sectionlinks_fontcolorhover, 1
                                                  lcl_scripts = lcl_scripts & "changePreviewColor('sectionlinks_fontcolorhover_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Link Row - View All URL
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" valign=""top"">View ALL URL:<br />(mouseover)</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
  response.write "                                <fieldset class=""fieldset"">" & vbcrlf
  response.write "                               <select id=""sectionlinks_viewall_urltype_" & lcl_rowType & "_" & p_rowcount & """ name=""sectionlinks_viewall_urltype_" & lcl_rowType & "_" & p_rowcount & """ style=""margin-bottom: 4px;"" onchange=""enableDisableViewALLURL('" & lcl_rowType & "_" & p_rowcount & "','Y');"">" & vbcrlf
                                                   displayCommunityLinkOptions "VIEWALL_URLTYPE", p_viewall_urltype
                                                   'lcl_scripts = lcl_scripts & "enableDisableViewALLURL('" & lcl_rowType & "_" & p_rowcount & "','Y');" & vbcrlf
  response.write "                               </select><br />" & vbcrlf
  response.write "                               <input type=""text"" id=""sectionlinks_viewall_url_" & lcl_rowType & "_" & p_rowcount & """ name=""sectionlinks_viewall_url_" & lcl_rowType & "_" & p_rowcount & """ class=""sectionlinks_viewall_url"" value=""" & p_viewall_url & """ /><br />" & vbcrlf
  response.write "                               <select id=""sectionlinks_viewall_url_wintype_" & lcl_rowType & "_" & p_rowcount & """ name=""sectionlinks_viewall_url_wintype_" & lcl_rowType & "_" & p_rowcount & """>" & vbcrlf
                                                   displayCommunityLinkOptions "VIEWALL_URL_WINTYPE", p_viewall_url_wintype
  response.write "                               </select>" & vbcrlf
  response.write "                                </fieldset>" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
 'END: Link Row Options -------------------------------------------------------

  response.write "                      </table>" & vbcrlf
  response.write "                      </fieldset>" & vbcrlf
  response.write "                  </td>" & vbcrlf

end sub
%>
