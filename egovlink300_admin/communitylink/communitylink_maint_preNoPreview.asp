<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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

'Check to see if any Mayor's Blog images exist and if so resize the borders around the image.
 lcl_onload = lcl_onload & "resizeBlogImgBorders();"

'Check for a CommunityLink record for the org.
'If one DOES exist then pull all of the values.
'If one does NOT exist then create it and enter the default values.
 lcl_communitylinkid = getCommunityLinkID(session("orgid"), session("userid"))

'Retrieve the CommunityLink record.
 getCommunityLinkInfo lcl_communitylinkid, session("orgid"), lcl_isEgovHomePage, lcl_website_size, lcl_website_size_customsize, _
                      lcl_website_alignment, lcl_website_bgcolor, lcl_showlogo, lcl_logo_filename, lcl_logo_filenamebg, _
                      lcl_logo_alignment, lcl_showtopbar, lcl_topbar_bgcolor, lcl_topbar_fonttype, lcl_topbar_fontcolor, _
                      lcl_topbar_fontcolorhover, lcl_showsidemenubar, lcl_sidemenubar_alignment, lcl_sidemenuoption_bgcolor, _
                      lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_sidemenuoption_fonttype, _
                      lcl_sidemenuoption_fontcolor, lcl_sidemenuoption_fontcolorhover, lcl_pageheader_alignment, _
                      lcl_pageheader_fontcolor, lcl_pageheader_fonttype, lcl_pageheader_bgcolor, lcl_showRSS, lcl_url_twitter, _
                      lcl_url_facebook, lcl_url_myspace, lcl_url_blogger

'Check for org features
 lcl_orghasfeature_administrationlink = orghasfeature("AdministrationLink")
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />
<!--  <link rel="stylesheet" type="text/css" href="../custom/css/dragdrop.css" /> -->

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>
<!--  <script language="javascript" src="../scripts/drag_drop.js"></script> -->
  <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<%
  response.write "<style type=""text/css"">" & vbcrlf
  response.write "  .orgLogo { padding: 5px; }" & vbcrlf

 'BEGIN: Top Bar Styles -------------------------------------------------------
  response.write "  .topBar " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      padding:          5px;" & vbcrlf
  response.write "      background-color: #" & lcl_topbar_bgcolor   & ";" & vbcrlf
  response.write "      color:            #" & lcl_topbar_fontcolor & ";" & vbcrlf
  response.write "    }" & vbcrlf

  response.write "  .topBarOption:link, .topBarOption:visited " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-family: "  & lcl_topbar_fonttype  & ";" & vbcrlf
  response.write "      color:       #" & lcl_topbar_fontcolor & ";" & vbcrlf
  response.write "      font-size:   10px;" & vbcrlf
  response.write "    }" & vbcrlf

  response.write "  .topBarOption:hover " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-family:      " & lcl_topbar_fonttype       & ";" & vbcrlf
  response.write "      color:           #" & lcl_topbar_fontcolorhover & ";" & vbcrlf
  response.write "      font-size:       10px;" & vbcrlf
  response.write "      text-decoration: underline;" & vbcrlf
  response.write "    }" & vbcrlf
 'END: Top Bar Styles ---------------------------------------------------------

 'BEGIN: Side Menubar Styles --------------------------------------------------
  if lcl_showsidemenubar then
     response.write "  .sideMenuBar" & vbcrlf
     response.write "    { " & vbcrlf
     response.write "      padding:          5px;" & vbcrlf
     response.write "      cursor:           pointer;" & vbcrlf
     response.write "      border-bottom:    1pt solid #ffffff;" & vbcrlf
     response.write "      background-color: #" & lcl_sidemenuoption_bgcolor & ";" & vbcrlf
     'response.write "      color: #" & lcl_topbarfontcolor & ";" & vbcrlf
     response.write "    }" & vbcrlf

     response.write "  .sideMenuBarOption:link, .sideMenuBarOption:visited " & vbcrlf
     response.write "    { " & vbcrlf
     response.write "      font-size:   12px;" & vbcrlf
     response.write "      font-family:  " & lcl_sidemenuoption_fonttype  & ";" & vbcrlf
     response.write "      color:       #" & lcl_sidemenuoption_fontcolor & ";" & vbcrlf
     response.write "    }" & vbcrlf

     response.write "  .sideMenuBarOption:hover " & vbcrlf
     response.write "    { " & vbcrlf
     response.write "      font-size: 12px;" & vbcrlf
     response.write "      font-family:  " & lcl_sidemenuoption_fonttype       & ";" & vbcrlf
     response.write "      color:       #" & lcl_sidemenuoption_fontcolorhover & ";" & vbcrlf
     response.write "      text-decoration: underline;" & vbcrlf
     response.write "    }" & vbcrlf
  end if
 'END: Side Menubar Styles ----------------------------------------------------

 'BEGIN: Page Header ----------------------------------------------------------
  response.write "  .pageHeader "
  response.write "    { " & vbcrlf
  response.write "      background-color: #" & lcl_pageheader_bgcolor & "; " & vbcrlf
  response.write "      border-bottom: 1pt solid #808080;" & vbcrlf
  response.write "      padding: 5px;" & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .pageHeader_welcome " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "     	font-family:    " & lcl_pageheader_fonttype  & "; " & vbcrlf
  response.write "     	color:         #" & lcl_pageheader_fontcolor & "; " & vbcrlf
  response.write "     	font-size:     12px; " & vbcrlf
  response.write "     	font-weight:   bold; " & vbcrlf
  response.write "     	padding:       0; " & vbcrlf
  response.write "     	margin-bottom: 2px; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .pageHeader_welcomeSubMsg " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "     	font-family:  " & lcl_pageheader_fonttype  & "; " & vbcrlf
  response.write "     	color:       #" & lcl_pageheader_fontcolor & "; " & vbcrlf
  response.write "      font-size:   10px; " & vbcrlf
  response.write "      font-weight: bold; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .pageHeader_homePageMsg " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "     	font-family:  " & lcl_pageheader_fonttype  & "; " & vbcrlf
  response.write "     	color:       #" & lcl_pageheader_fontcolor & "; " & vbcrlf
  response.write "     	font-size:   12px; " & vbcrlf
  response.write "      padding-top: 5px; " & vbcrlf
  response.write "    } " & vbcrlf
 'END: Page Header ------------------------------------------------------------

 'BEGIN: Community Link Options -----------------------------------------------
  response.write "  .communitylink_table " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      border: 1pt solid #000000; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communitylink_table th " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      font-weight:      bold; " & vbcrlf
  response.write "      border-bottom:    1pt solid #000000; " & vbcrlf
  response.write "      background-color: #cccccc; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communityLink_viewLinks:link, .communityLink_viewLinks:visited, .communityLink_viewLinks:active " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      color: #800000; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communityLink_viewLinks:hover " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      color: #800000; " & vbcrlf
  response.write "      text-decoration: underline; " & vbcrlf
  response.write "    } " & vbcrlf

  response.write "  .communityLink_bottomborder " & vbcrlf
  response.write "    { " & vbcrlf
  response.write "      border-bottom: 1pt dotted #808080; "  & vbcrlf
  response.write "    } " & vbcrlf
 'END: Community Link Options -------------------------------------------------

  response.write "</style>" & vbcrlf

 'BEGIN: Javascripts ----------------------------------------------------------
  response.write "<script type=""text/javascript"" src=""https://s7.addthis.com/js/200/addthis_widget.js""></script>" & vbcrlf
%>
<script language="javascript">
<!--
//Variable for AddThis button
var addthis_pub="cschappacher";

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

  //Background Logo
		if (document.getElementById("logo_filenamebg").value!="") {
      lcl_logofilenamebg = document.getElementById("logo_filenamebg").value.toUpperCase();
      lcl_ext_start_pos  = lcl_logofilenamebg.indexOf(".");
      lcl_ext            = lcl_logofilenamebg.substr(lcl_ext_start_pos+1,lcl_logofilenamebg.length);

      if(<%=lcl_imgTypes%>) {
         clearMsg("findImageButtonbg");
     }else{
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
        document.getElementById("website_size_customsize").focus();
        inlineMsg(document.getElementById("website_size_customsize").id,'<strong>Required Field Missing: </strong> Website Size (Custom Pixel Size)',10,'website_size_customsize');
        lcl_false_count = lcl_false_count + 1;
     }else{
        var rege = /^\d+$/;
        var Ok   = rege.exec(document.getElementById("website_size_customsize").value);

     			if (! Ok) {
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
     document.getElementById("communitylink_maint").submit();
     return true;
  }
}

function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) {
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

  if(lcl_folderStart > 0) {
     lcl_showFolderStart = "&folderStart=unpublished_documents";
  }

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += "&returnAsHTMLLink=N";
  pickerURL += lcl_showFolderStart;
  pickerURL += lcl_displayDocuments;
  pickerURL += lcl_displayActionLine;
  pickerURL += lcl_displayPayments;
  pickerURL += lcl_displayURL;

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
      textEl.value  = text;
}

function setupMenuOption(iType,iRowID) {
  if(iType=="OVER") {
     lcl_optionbg      = '<%=lcl_sidemenuoption_bgcolorhover%>';
     lcl_textcolor     = '<%=lcl_sidemenuoption_fontcolorhover%>';
     lcl_showUnderLine = 'underline';
  }else{
     lcl_optionbg      = '<%=lcl_sidemenuoption_bgcolor%>';
     lcl_textcolor     = '<%=lcl_sidemenuoption_fontcolor%>';
     lcl_showUnderLine = 'none';
  }

  document.getElementById("sideMenuBar"       + iRowID).style.backgroundColor="#" + lcl_optionbg;
  document.getElementById("sideMenuBarOption" + iRowID).style.backgroundColor="#" + lcl_optionbg;
  document.getElementById("sideMenuBarOption" + iRowID).style.color="#"           + lcl_textcolor;
  document.getElementById("sideMenuBarOption" + iRowID).style.textDecoration = lcl_showUnderLine;
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

  document.getElementById(iFieldID+"_previewcolor").style.backgroundColor=lcl_color;
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
     document.getElementById("styleProperties_Portal_"       + iRowCount).style.display = lcl_displayed;
     document.getElementById("styleProperties_DisplayOrder_" + iRowCount).style.display = lcl_displayed;
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

function resizeBlogImgBorders() {
  lcl_total_blogimgs = document.getElementsByName("blogimg").length;

  for (i=1; i<=lcl_total_blogimgs; ++ i) {
       if(document.getElementById("blogimg_"+i)) {
          //Get the height of the blog image.
          lcl_total_height = document.getElementById("blogimg_"+i).height;

          //Get the widths of the blog images.
          lcl_blogimgleft_width  = document.getElementById("blogimg_left_"+i).width;
          lcl_blogimgright_width = document.getElementById("blogimg_right_"+i).width;
          lcl_blogimg_width      = document.getElementById("blogimg_"+i).width;
          lcl_total_width        = lcl_blogimgleft_width + lcl_blogimgright_width + lcl_blogimg_width;

          //Adjust the top/bottom blog images with the proper width.
          document.getElementById("blogimg_top_"+i).width    = lcl_total_width;
          document.getElementById("blogimg_bottom_"+i).width = lcl_total_width;

          //Adjust the left/right blog images with the proper height.
          document.getElementById("blogimg_left_"+i).height  = lcl_total_height
          document.getElementById("blogimg_right_"+i).height = lcl_total_height;
       }
  }
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
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<!--<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<% 'lcl_onload%>"> -->

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 
<%
  lcl_labelcolumn_width = "150"

  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<table border=""0"" bordercolor=""#00000ff"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""800"">" & vbcrlf
  response.write "  <form name=""communitylink_maint"" id=""communitylink_maint"" action=""communitylink_action.asp"" method=""post"">" & vbcrlf
  response.write "    <input type=""hidden"" name=""communitylinkid"" id=""communitylinkid"" value=""" & lcl_communitylinkid & """ size=""5"" maxlength=""10"" />" & vbcrlf

 'BEGIN: Page Title and Screen Messages ---------------------------------------
  response.write "  <caption align=""left"">" & vbcrlf
  response.write "    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "      <tr>" & vbcrlf
  response.write "          <td><font size=""+1""><strong>" & session("sOrgName") & "&nbsp;" & lcl_pagetitle & "</strong></font></td>" & vbcrlf
  response.write "          <td align=""right""><span id=""screenMsg"" style=""color:#ff0000; font-size:10pt; font-weight:bold;""></span></td>" & vbcrlf
  response.write "      </tr>" & vbcrlf
  response.write "    </table>" & vbcrlf

  displayButtons "MAINT"

  response.write "  </caption>" & vbcrlf
 'END: Page Title and Screen Messages -----------------------------------------

 'BEGIN: Layout Options ------------------------------------------------------
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"" colspan=""2"">" & vbcrlf
  response.write "          <fieldset>" & vbcrlf
  response.write "            <legend>Layout Options&nbsp;</legend>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Website Size
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Size:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""website_size"" id=""website_size"" onchange=""setupCustomSize();"">" & vbcrlf

  displayCommunityLinkOptions "WEBSITE_SIZE", lcl_website_size

  response.write "                      </select>&nbsp;" & vbcrlf
  response.write "                      <span id=""website_size_customsize_span"">" & vbcrlf
  response.write "                      <input type=""text"" name=""website_size_customsize"" id=""website_size_customsize"" value=""" & lcl_website_size_customsize & """ size=""5"" maxlength=""10"" onchange=""clearMsg('website_size_customsize');"" />&nbsp;" & vbcrlf
  response.write "                      <font style=""font-size:10px; color:#800000"">(All sizes in pixels)</font></span>" & vbcrlf
  response.write "                  </td>" & vbcrlf

 'Is Egov Home Page
  if lcl_isEgovHomePage then
     lcl_checked_egovhome = " checked=""checked"""
  else
     lcl_checked_egovhome = ""
  end if

  response.write "                  <td align=""right"">" & vbcrlf
  response.write "                      Make CommunityLink E-Gov Home Page:" & vbcrlf
  response.write "                      <input type=""checkbox"" name=""isEgovHomePage"" id=""isEgovHomePage"" value=""on""" & lcl_checked_egovhome & " />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Website Alignment
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Alignment:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <select name=""website_alignment"" id=""website_alignment"">" & vbcrlf

  displayCommunityLinkOptions "WEBSITE_ALIGN", lcl_website_alignment

  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Website Background Color
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Background Color:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
                                        setupColorSelection "website_bgcolor", lcl_website_bgcolor, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('website_bgcolor');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  response.write "              <tr><td colspan=""3"">&nbsp;</td></tr>" & vbcrlf

 'Show Logo
  if lcl_showlogo then
     lcl_checked_logo = " checked=""checked"""
  else
     lcl_checked_logo = ""
  end if

  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Show Logo:</td>" & vbcrlf
  response.write "                  <td colspan=""2""><input type=""checkbox"" name=""showlogo"" id=""showlogo"" value=""on""" & lcl_checked_logo & " /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Logo Alignment
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Alignment:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <select name=""logo_alignment"" id=""logo_alignment"">" & vbcrlf

  displayCommunityLinkOptions "WEBSITE_LOGO_ALIGN", lcl_logo_alignment

  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Logo Filename
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Logo:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""input"" name=""logo_filename"" id=""logo_filename"" value=""" & lcl_logo_filename & """ size=""50"" maxlength=""500"" onchange=""clearMsg('findImageButton');"" />&nbsp;" & vbcrlf
  response.write "                      <input type=""button"" name=""findImageButton"" id=""findImageButton"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButton');doPicker('communitylink_maint.logo_filename','Y','','','');"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Logo Filename - Background
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Background Logo:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""input"" name=""logo_filenamebg"" id=""logo_filenamebg"" value=""" & lcl_logo_filenamebg & """ size=""50"" maxlength=""500"" onchange=""clearMsg('findImageButtonbg');"" />&nbsp;" & vbcrlf
  response.write "                      <input type=""button"" name=""findImageButtonbg"" id=""findImageButtonbg"" value=""Find Image"" class=""button"" onclick=""clearMsg('findImageButtonbg');doPicker('communitylink_maint.logo_filenamebg','Y','','','');"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  response.write "              <tr><td colspan=""3"">&nbsp;</td></tr>" & vbcrlf

 'Show RSS
  if lcl_showRSS then
     lcl_checked_showRSS = " checked=""checked"""
  else
     lcl_checked_showRSS = ""
  end if

  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Show RSS (icon):</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""checkbox"" name=""showRSS"" id=""showRSS"" value=""on""" & lcl_checked_showRSS & " />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'URL - Twitter
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>URL - Twitter:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""input"" name=""url_twitter"" id=""url_twitter"" value=""" & lcl_url_twitter & """ size=""50"" maxlength=""500"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'URL - Facebook
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>URL - Facebook:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""input"" name=""url_facebook"" id=""url_facebook"" value=""" & lcl_url_facebook & """ size=""50"" maxlength=""500"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'URL - MySpace
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>URL - MySpace:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""input"" name=""url_myspace"" id=""url_myspace"" value=""" & lcl_url_myspace & """ size=""50"" maxlength=""500"" />" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'URL - Blogger
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>URL - Blogger:</td>" & vbcrlf
  response.write "                  <td colspan=""2"">" & vbcrlf
  response.write "                      <input type=""input"" name=""url_blogger"" id=""url_blogger"" value=""" & lcl_url_blogger & """ size=""50"" maxlength=""500"" />&nbsp;" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
 'END: Website Logo -----------------------------------------------------------

  response.write "  <tr>" & vbcrlf

  response.write "      <td valign=""top"">" & vbcrlf

 'BEGIN: Top Bar/Footer Options -----------------------------------------------
  response.write "          <fieldset>" & vbcrlf
  response.write "            <legend>Top Bar/Footer Options&nbsp;</legend>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Top Bar
  if lcl_showtopbar then
     lcl_checked_topbar = " checked=""checked"""
  else
     lcl_checked_topbar = ""
  end if

  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Show Top Bar:</td>" & vbcrlf
  response.write "                  <td><input type=""checkbox"" name=""showtopbar"" id=""showtopbar"" value=""on""" & lcl_checked_topbar & " /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Background Color
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Background Color:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "topbar_bgcolor", lcl_topbar_bgcolor, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('topbar_bgcolor');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Font Type
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Font Type:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""topbar_fonttype"" id=""topbar_fonttype"">" & vbcrlf

  displayCommunityLinkOptions "TOPBAR_FONTTYPE", lcl_topbar_fonttype

  response.write "                      </select>" & vbcrf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Font Color
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Font Color:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "topbar_fontcolor", lcl_topbar_fontcolor, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('topbar_fontcolor');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Font Color - Hover
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Font Color<br />(mouseover):</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "topbar_fontcolorhover", lcl_topbar_fontcolorhover, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('topbar_fontcolorhover');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
 'END: Top Bar/Footer Options -------------------------------------------------

 'BEGIN: Page Header Options --------------------------------------------------
  response.write "          <fieldset>" & vbcrlf
  response.write "            <legend>Page Header Options&nbsp;</legend>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Background Color
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Background Color:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "pageheader_bgcolor", lcl_pageheader_bgcolor, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('pageheader_bgcolor');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Page Header Alignment
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Alignment:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""pageheader_alignment"" id=""pageheader_alignment"">" & vbcrlf

  displayCommunityLinkOptions "PAGEHEADER_ALIGN", lcl_pageheader_alignment

  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Page Header Color
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Font Color:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "pageheader_fontcolor", lcl_pageheader_fontcolor, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('pageheader_fontcolor');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Page Header - Font Type
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Font Type:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""pageheader_fonttype"" id=""pageheader_fonttype"">" & vbcrlf

  displayCommunityLinkOptions "PAGEHEADER_FONTTYPE", lcl_pageheader_fonttype

  response.write "                      </select>" & vbcrf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
 'END: Page Header Options ----------------------------------------------------

  response.write "      </td>" & vbcrlf

 'BEGIN: Side Menu Bar Options ------------------------------------------------
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <fieldset>" & vbcrlf
  response.write "            <legend>Side Menu Bar Options&nbsp;</legend>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf

 'Show Side Menubar
  if lcl_showsidemenubar then
     lcl_checked_sidemenubar = " checked=""checked"""
  else
     lcl_checked_sidemenubar = ""
  end if

  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Show Side Menu Bar:</td>" & vbcrlf
  response.write "                  <td><input type=""checkbox"" name=""showsidemenubar"" id=""showsidemenubar"" value=""on""" & lcl_checked_sidemenubar & " /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Side Menubar Alignment
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Alignment:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""sidemenubar_alignment"" id=""sidemenubar_alignment"">" & vbcrlf

  displayCommunityLinkOptions "SIDEMENUBAR_ALIGN", lcl_sidemenubar_alignment

  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Side Menubar Option Color
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Option Color:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "sidemenuoption_bgcolor", lcl_sidemenuoption_bgcolor, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_bgcolor');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Side Menubar Option Color - Hover
  response.write "              <tr>" & vbcrlf
  response.write "                  <td width=""" & lcl_labelcolumn_width & """>Option Color<br />(mouseover):</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "sidemenuoption_bgcolorhover", lcl_sidemenuoption_bgcolorhover, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_bgcolorhover');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  response.write "              <tr><td colspan=""2"">&nbsp;</td></tr>" & vbcrlf

 'Side Menubar Option Alignment
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Text Alignment:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""sidemenuoption_alignment"" id=""sidemenuoption_alignment"">" & vbcrlf

  displayCommunityLinkOptions "SIDEMENUOPT_TEXTALIGN", lcl_sidemenuoption_alignment

  response.write "                      </select>" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Side Menubar Option - Font Type
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Option Font Type:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
  response.write "                      <select name=""sidemenuoption_fonttype"" id=""sidemenuoption_fonttype"">" & vbcrlf

  displayCommunityLinkOptions "SIDEMENUOPT_FONTTYPE", lcl_sidemenuoption_fonttype

  response.write "                      </select>" & vbcrf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Side Menubar Option - Font Color
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Option Font Color:</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "sidemenuoption_fontcolor", lcl_sidemenuoption_fontcolor, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_fontcolor');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Side Menubar Option - Font Color - Hover
  response.write "              <tr>" & vbcrlf
  response.write "                  <td>Option Font Color<br />(mouseover):</td>" & vbcrlf
  response.write "                  <td>" & vbcrlf
                                        setupColorSelection "sidemenuoption_fontcolorhover", lcl_sidemenuoption_fontcolorhover, 1
                                        lcl_scripts = lcl_scripts & "changePreviewColor('sidemenuoption_fontcolorhover');" & vbcrlf
  response.write "                  </td>" & vbcrlf
  response.write "              </tr>" & vbcrlf

  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
 'END: Side Menu Bar Options --------------------------------------------------

  response.write "  </tr>" & vbcrlf

 'BEGIN: CommunityLink Options ------------------------------------------------
  iBGColor              = "#eeeeee"
  iRowCount             = 0
  lcl_features_shown    = ""
  iTotalCLFeaturesAvail = getCLFeatAvailCount(session("orgid"))

  response.write "  <tr>" & vbcrlf
  response.write "      <td colspan=""2"" valign=""top"">" & vbcrlf
  response.write "          <fieldset>" & vbcrlf
  response.write "            <legend>CommunityLink Options&nbsp;</legend>" & vbcrlf
  response.write "            <p>" & vbcrlf
  response.write "            <table border=""0"" cellspacing=""0"" cellpadding=""3"" class=""communitylink_table"">" & vbcrlf
  response.write "              <tr>" & vbcrlf
  response.write "                  <th align=""left"">CommunityLink Sections</th>" & vbcrlf
  response.write "                  <th nowrap=""nowrap"">Community Link</th>" & vbcrlf
  response.write "                  <th nowrap=""nowrap"">Savvy/IFRAME</th>" & vbcrlf
  response.write "              </tr>" & vbcrlf

 'Setup the defaults.
  lcl_isCommunityLinkOn       = 0
  lcl_isSavvyOn               = 0
  lcl_sectionheader_bgcolor   = "ffffff"
  lcl_sectionheader_linecolor = "000000"
  lcl_sectionheader_fonttype  = getCLOptionDefault("SECTIONHEADER_FONTTYPE")
  lcl_sectionheader_fontcolor = "000000"
  lcl_sectiontext_bgcolor     = "ffffff"
  lcl_sectiontext_fonttype    = getCLOptionDefault("SECTIONTEXT_FONTTYPE")
  lcl_sectiontext_fontcolor   = "000000"
  lcl_sectionlinks_alignment  = getCLOptionDefault("SECTIONLINKS_ALIGN")
  lcl_sectionlinks_fonttype   = getCLOptionDefault("SECTIONLINKS_FONTTYPE")
  lcl_sectionlinks_fontcolor  = "800000"

  sSQL = " SELECT f.featureid, "
  sSQL = sSQL & " isnull(f.CL_portaltype,'') AS portaltype, "
  sSQL = sSQL & " f.CommunityLinkOn, "
  sSQL = sSQL & " isnull(cl.featurename, isnull(otf.featurename, f.featurename)) AS featurename, "
  sSQL = sSQL & " isnull(otf.featurename, f.featurename) AS featurename_original, "
  sSQL = sSQL & " isnull(cl.portalcolumn, 1) AS portalcolumn, "
  sSQL = sSQL & " isnull(cl.displayorder, 1) AS displayorder, "
  sSQL = sSQL & " isnull(cl.numListItemsShown_CL, f.CL_numListItems) AS numListItemsShown_CL, "
  sSQL = sSQL & " isnull(cl.numListItemsShown_SAVVY, f.CL_numListItems) AS numListItemsShown_SAVVY, "
  sSQL = sSQL & " f.CL_numListItems AS numListItemsShown_original, "
  sSQL = sSQL & " isnull(cl.isCommunityLinkOn,"                  & lcl_isCommunityLinkOn       & ") AS isCommunityLinkOn, "
  sSQL = sSQL & " isnull(cl.isSavvyOn,"                          & lcl_isSavvyOn               & ") AS isSavvyOn, "
  sSQL = sSQL & " isnull(cl.sectionheader_bgcolor_CL,'"          & lcl_sectionheader_bgcolor   & "') AS sectionheader_bgcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_bgcolor_Savvy,'"       & lcl_sectionheader_bgcolor   & "') AS sectionheader_bgcolor_Savvy, "
  sSQL = sSQL & " isnull(cl.sectionheader_linecolor_CL,'"        & lcl_sectionheader_linecolor & "') AS sectionheader_linecolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_linecolor_Savvy,'"     & lcl_sectionheader_linecolor & "') AS sectionheader_linecolor_Savvy, "
  sSQL = sSQL & " isnull(cl.sectionheader_fonttype_CL,'"         & lcl_sectionheader_fonttype  & "') AS sectionheader_fonttype_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_fonttype_Savvy,'"      & lcl_sectionheader_fonttype  & "') AS sectionheader_fonttype_Savvy, "
  sSQL = sSQL & " isnull(cl.sectionheader_fontcolor_CL,'"        & lcl_sectionheader_fontcolor & "') AS sectionheader_fontcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionheader_fontcolor_Savvy,'"     & lcl_sectionheader_fontcolor & "') AS sectionheader_fontcolor_Savvy, "
  sSQL = sSQL & " isnull(cl.sectiontext_bgcolor_CL,'"            & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_bgcolor_Savvy,'"         & lcl_sectiontext_bgcolor     & "') AS sectiontext_bgcolor_Savvy, "
  sSQL = sSQL & " isnull(cl.sectiontext_fonttype_CL,'"           & lcl_sectiontext_fonttype    & "') AS sectiontext_fonttype_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_fonttype_Savvy,'"        & lcl_sectiontext_fonttype    & "') AS sectiontext_fonttype_Savvy, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontcolor_CL,'"          & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectiontext_fontcolor_Savvy,'"       & lcl_sectiontext_fontcolor   & "') AS sectiontext_fontcolor_Savvy, "
  sSQL = sSQL & " isnull(cl.sectionlinks_alignment_CL,'"         & lcl_sectionlinks_alignment  & "') AS sectionlinks_alignment_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_alignment_Savvy,'"      & lcl_sectionlinks_alignment  & "') AS sectionlinks_alignment_Savvy, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fonttype_CL,'"          & lcl_sectionlinks_fonttype   & "') AS sectionlinks_fonttype_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fonttype_Savvy,'"       & lcl_sectionlinks_fonttype   & "') AS sectionlinks_fonttype_Savvy, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolor_CL,'"         & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolor_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolor_Savvy,'"      & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolor_Savvy, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolorhover_CL,'"    & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolorhover_CL, "
  sSQL = sSQL & " isnull(cl.sectionlinks_fontcolorhover_Savvy,'" & lcl_sectionlinks_fontcolor  & "') AS sectionlinks_fontcolorhover_Savvy "
  sSQL = sSQL & " FROM egov_communitylink_displayorgfeatures cl "
  sSQL = sSQL &      " RIGHT OUTER JOIN egov_organizations_to_features otf "
  sSQL = sSQL &      " INNER JOIN egov_organization_features f "
  sSQL = sSQL &      " ON otf.featureid = f.featureid "
  sSQL = sSQL &      " ON f.featureid = cl.featureid "
  sSQL = sSQL &      " AND cl.orgid = otf.orgid "
  sSQL = sSQL & " WHERE f.haspublicview = 1 "
  sSQL = sSQL & " AND f.CommunityLinkOn = 1 "
  sSQL = sSQL & " AND otf.orgid = " & session("orgid")
  sSQL = sSQL & " ORDER BY cl.isCommunityLinkOn DESC, isnull(cl.portalcolumn, 0), isnull(cl.displayorder, 0), "
  sSQL = sSQL &          " isnull(cl.featurename, isnull(otf.featurename, f.featurename)), cl.isSavvyOn DESC "

  set oCLFeatures = Server.CreateObject("ADODB.Recordset")
  oCLFeatures.Open sSQL, Application("DSN"), 3, 1

  if not oCLFeatures.eof then
     do while not oCLFeatures.eof
        iBGColor  = changeBGColor(iBGColor,"#eeeeee","#ffffff")
        iRowCount = iRowCount + 1

        response.write "              <tr valign=""top"" bgcolor=""" & iBGColor & """>" & vbcrlf

        displayCommunityLinkFeatureOptions "CL", iRowCount, iTotalCLFeaturesAvail, iBGColor, lcl_scripts, _
                                           oCLFeatures("featureid"), _
                                           oCLFeatures("featurename"), _
                                           oCLFeatures("featurename_original"), _
                                           oCLFeatures("portalcolumn"), _
                                           oCLFeatures("displayorder"), _
                                           oCLFeatures("numListItemsShown_CL"), _
                                           oCLFeatures("numListItemsShown_original"), _
                                           oCLFeatures("isCommunityLinkOn"), _
                                           oCLFeatures("sectionheader_bgcolor_CL"), _
                                           oCLFeatures("sectionheader_linecolor_CL"), _
                                           oCLFeatures("sectionheader_fonttype_CL"), _
                                           oCLFeatures("sectionheader_fontcolor_CL"), _
                                           oCLFeatures("sectiontext_bgcolor_CL"), _
                                           oCLFeatures("sectiontext_fonttype_CL"), _
                                           oCLFeatures("sectiontext_fontcolor_CL"), _
                                           oCLFeatures("sectionlinks_alignment_CL"), _
                                           oCLFeatures("sectionlinks_fonttype_CL"), _
                                           oCLFeatures("sectionlinks_fontcolor_CL"), _
                                           oCLFeatures("sectionlinks_fontcolorhover_CL")

        displayCommunityLinkFeatureOptions "SAVVY", iRowCount, iTotalCLFeaturesAvail, iBGColor, lcl_scripts, _
                                           oCLFeatures("featureid"), _
                                           oCLFeatures("featurename"), _
                                           oCLFeatures("featurename_original"), _
                                           oCLFeatures("portalcolumn"), _
                                           oCLFeatures("displayorder"), _
                                           oCLFeatures("numListItemsShown_SAVVY"), _
                                           oCLFeatures("numListItemsShown_original"), _
                                           oCLFeatures("isSavvyOn"), _
                                           oCLFeatures("sectionheader_bgcolor_Savvy"), _
                                           oCLFeatures("sectionheader_linecolor_Savvy"), _
                                           oCLFeatures("sectionheader_fonttype_Savvy"), _
                                           oCLFeatures("sectionheader_fontcolor_Savvy"), _
                                           oCLFeatures("sectiontext_bgcolor_Savvy"), _
                                           oCLFeatures("sectiontext_fonttype_Savvy"), _
                                           oCLFeatures("sectiontext_fontcolor_Savvy"), _
                                           oCLFeatures("sectionlinks_alignment_Savvy"), _
                                           oCLFeatures("sectionlinks_fonttype_Savvy"), _
                                           oCLFeatures("sectionlinks_fontcolor_Savvy"), _
                                           oCLFeatures("sectionlinks_fontcolorhover_Savvy")

        response.write "              </tr>" & vbcrlf

        oCLFeatures.movenext
     loop
  end if

  oCLFeatures.close
  set oCLFeatures = nothing

 'Total Rows
  response.write "              <tr>" & vbcrlf
  response.write "                  <td><input type=""hidden"" name=""totalCLRows"" id=""totalCLRows"" value=""" & iRowCount & """ size=""3"" maxlength=""10"" /></td>" & vbcrlf
  response.write "              </tr>" & vbcrlf
  response.write "            </table>" & vbcrlf
  response.write "            </p>" & vbcrlf
  response.write "          </fieldset>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
 'END: CommunityLink Options --------------------------------------------------

  response.write "  </form>" & vbcrlf
  response.write "</table>" & vbcrlf

  displayButtons "MAINT"

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf

 '-----------------------------------------------------------------------------
 'BEGIN: CommunityLink Preview
 '-----------------------------------------------------------------------------
  lcl_website_width = getWebsiteWidth(lcl_website_size, lcl_website_size_customsize)

  response.write "<fieldset>" & vbcrlf
  response.write "  <legend><a name=""communitylink_preview"">CommunityLink Preview</a>&nbsp;</legend>" & vbcrlf
  response.write "<p>" & vbcrlf

  displayButtons "PREVIEW"

  response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""border:1pt solid #000000;"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"" align=""" & lcl_website_alignment & """ bgcolor=""" & lcl_website_bgcolor & """>" & vbcrlf
  response.write "          <table border=""0"" bordercolor=""#ff0000"" cellspacing=""0"" cellpadding=""0"" width=""" & lcl_website_width & """ bgcolor=""#ffffff"" style=""border:1pt solid #000000;"">" & vbcrlf

 'BEGIN: Show Logo ------------------------------------------------------------
  if lcl_showlogo then
    'Build the Logo URLs
     lcl_orgLogoURL = session("egovclientwebsiteurl")
     lcl_orgLogoURL = lcl_orgLogoURL & "/admin/custom/pub/"
     lcl_orgLogoURL = lcl_orgLogoURL & session("virtualdirectory")
     lcl_orgLogoURL = lcl_orgLogoURL & "/unpublished_documents"

     if lcl_logo_filename <> "" then
        lcl_logo_filename = lcl_orgLogoURL & lcl_logo_filename
     else
        lcl_logo_filename = getDefaultLogo("LEFT",session("orgid"))
     end if

     if lcl_logo_filenamebg <> "" then
        lcl_logo_filenamebg = lcl_orgLogoURL & lcl_logo_filenamebg
     else
        lcl_logo_filenamebg = getDefaultLogo("RIGHT",session("orgid"))
     end if

     lcl_orgLogo = "<img src=""" & lcl_logo_filename & """ name=""orgLogo"" id=""orgLogo"" />"

    'If the logofilenamebg is NULL then display the logo bgcolor
     if lcl_logo_filenamebg <> "" then
        lcl_orgLogoBGstyle = "background-image:url('" & lcl_logo_filenamebg & "');"
     end if

     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""3"" class=""orgLogo"" align=""" & lcl_logo_alignment & """>" & vbcrlf
     response.write "                    <div style=""" & lcl_orgLogoBGstyle & """>" & lcl_orgLogo & "</div>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: Show Logo --------------------------------------------------------------

 'BEGIN: Show TopBar ----------------------------------------------------------
  if lcl_showtopbar then
     sUserName = ", <strong>" & GetAdminName(session("userid")) & "</strong>"

     response.write "            <tr>" & vbcrlf
     response.write "                <td colspan=""3"" class=""topBar"">" & vbcrlf
     response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""topBar"">" & vbcrlf
     response.write "                      <tr>" & vbcrlf
     response.write "                          <td align=""left"">" & vbcrlf
  			response.write "                              <i>Today is " & FormatDateTime(Date(), vbLongDate) & ".</i>&nbsp;&nbsp;Welcome" & sUserName & "!" & vbcrlf
     response.write "                          </td>" & vbcrlf
     response.write "                          <td align=""right"">" & vbcrlf
                                                   ShowLoggedinLinks session("orgid")
     response.write "                          </td>" & vbcrlf
     response.write "                      </tr>" & vbcrlf
     response.write "                    </table>" & vbcrlf
     response.write "                </td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
  end if
 'END: Show TopBar ------------------------------------------------------------

  response.write "            <tr valign=""top"">" & vbcrlf

 'BEGIN: Build the column widths ----------------------------------------------
  if lcl_showsidemenubar then
     lcl_sidemenubar_width = 200
     lcl_pageheader_width  = lcl_website_width - lcl_sidemenubar_width
  else
     lcl_sidemenubar_width = 0
     lcl_pageheader_width  = lcl_website_width - lcl_sidemenubar_width
  end if

  lcl_column1_width = lcl_pageheader_width * 0.55
  lcl_column2_width = lcl_pageheader_width * 0.45
 'END: Build the column widths ------------------------------------------------

 'BEGIN: Side Menubar (LEFT) --------------------------------------------------
  if lcl_showsidemenubar AND lcl_sidemenubar_alignment = "LEFT" then
     response.write "                <td rowspan=""2"" nowrap=""nowrap"" style=""width:" & lcl_sidemenubar_width & "px; background-color:#" & lcl_sidemenuoption_bgcolor & """>" & vbcrlf

     displaySideMenubar session("orgid"), lcl_sidemenuoption_bgcolor, lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_isEgovHomePage

     response.write "                </td>" & vbcrlf
  end if
 'END: Side Menubar (LEFT) ----------------------------------------------------

 'BEGIN: Page Header ----------------------------------------------------------
  lcl_orgname       = getOrgName(session("orgid"))
  lcl_orgname_label = lcl_orgname

  if getState(session("orgid")) <> "" then
     lcl_orgname_label = lcl_orgname_label & ", " & getState(session("orgid"))
  end if

  lcl_tagline = getOrgTagLine(session("orgid"))

  if lcl_tagline <> "" then
     lcl_orgname_label = lcl_orgname_label & ", " & lcl_tagline
  end if

 'Find the length of the page header minus the AddThis button width
  lcl_pageheadertext_width = lcl_pageheader_width - 125

  response.write "                <td colspan=""2"" style=""width:" & lcl_pageheader_width & "px;"" align=""left"" class=""pageHeader"">" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"" class=""pageHeader_homePageMsg"">" & vbcrlf
  response.write "                      <tr valign=""top"">" & vbcrlf
  response.write "                          <td width=""" & lcl_pageheadertext_width & """ align=""" & lcl_pageheader_alignment & """>" & vbcrlf
  response.write "                              <div class=""pageHeader_welcome"">" & lcl_orgname_label & " - CommunityLink</div>" & vbcrlf
  response.write "                              <div class=""pageHeader_welcomeSubMsg"">Your connection to " & lcl_orgname & "</div><br />" & vbcrlf

 'Display the "page header" if the org has an "Edit Display" for the "homepage message".
  if orghasdisplay(session("orgid"),"homepage message") then
     response.write "                           <span class=""pageHeader_homePageMsg"">" & vbcrlf
	 			response.write                                getOrgDisplay(session("orgid"),"homepage message")
     response.write "                           </span>" & vbcrlf
  end if

  response.write "                          </td>" & vbcrlf
  response.write "                          <td align=""right"" style=""padding-right:5px;"">" & vbcrlf
                                                displayAddThisButton()
                                                getSocialSiteIcons "H", lcl_showRSS, lcl_url_twitter, lcl_url_facebook, _
                                                                   lcl_url_myspace, lcl_url_blogger

  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
 'END: Page Header ------------------------------------------------------------

 'BEGIN: Side Menubar (RIGHT) -------------------------------------------------
  if lcl_showsidemenubar AND lcl_sidemenubar_alignment = "RIGHT" then
     response.write "                <td rowspan=""2"" nowrap=""nowrap"" style=""width:" & lcl_sidemenubar_width & "px; background-color:#" & lcl_sidemenuoption_bgcolor & """>" & vbcrlf

     displaySideMenubar session("orgid"), lcl_sidemenuoption_bgcolor, lcl_sidemenuoption_bgcolorhover, lcl_sidemenuoption_alignment, lcl_isEgovHomePage

     response.write "                </td>" & vbcrlf
  end if

  response.write "            </tr>" & vbcrlf
 'END: Side Menubar (RIGHT) ---------------------------------------------------

 'BEGIN: CommunityLink Columns ------------------------------------------------
  response.write "            <tr valign=""top"">" & vbcrlf
                                  displayPortalSections "CL", 1, session("orgid"), True, 0, "Y", lcl_column1_width, "Y"
                                  displayPortalSections "CL", 2, session("orgid"), True, 0, "Y", lcl_column2_width, "Y"
  response.write "            </tr>" & vbcrlf
 'END: CommunityLink Columns --------------------------------------------------

 'BEGIN: Footer ---------------------------------------------------------------
  lcl_cityhome_label = GetOrgDisplay(session("orgid"),"homewebsitetag")

  if lcl_cityhome_label = "" then
     lcl_cityhome_label = "City Home"
  end if

  response.write "            <tr>" & vbcrlf
  response.write "                <td colspan=""3"" class=""topBar"">" & vbcrlf
  response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "                      <tr>" & vbcrlf
  response.write "                          <td>&nbsp;</td>" & vbcrlf
  response.write "                          <td align=""center"" class=""topBar"">" & vbcrlf
  response.write "                              <a href=""#communitylink_preview"" class=""topBarOption"">" & lcl_cityhome_label & "</a> |" & vbcrlf
  response.write "                              <a href=""#communitylink_preview"" class=""topBarOption"">E-Gov Home</a>" & vbcrlf
                                                ShowPublicDefaultFooterNav session("orgid"), 2, lcl_isEgovHomePage
  response.write "                              <br />" & vbcrlf

  if OrgHasDisplay(session("orgid"),"privacy policy") then
     response.write "                           <a href=""#communitylink_preview"" class=""topBarOption""><strong>Privacy Policy</strong></a> | " & vbcrlf
  end if

  if OrgHasDisplay(session("orgid"),"refund policy") then
     response.write "                           <a href=""#communitylink_preview"" class=""topBarOption"">Refund Policy</a> | " & vbcrlf
  end if

  response.write "                            		<a href=""#communitylink_preview"" class=""topBarOption"">Login</a> |" & vbcrlf
  response.write "                              <a href=""#communitylink_preview"" class=""topBarOption"">Register</a>" & vbcrlf
  response.write "                              <p>" & vbcrlf
  response.write "                              Copyright &copy; 2004-" & year(now) & " electronic commerce link, inc. dba <a href=""#communitylink_preview"" target=""_NEW"" class=""topBarOption"">egovlink</a>&nbsp;" & iDisplayTime & vbcrlf

 'BEGIN: DEMO CHECK TO ADD ADMIN LINK
  if lcl_orghasfeature_administrationlink then
     response.write "                           &nbsp;&nbsp;&nbsp;<a target=""_new"" href=""#communitylink_preview"" class=""topBarOption"">Administrator</a>" & vbcrlf
  end if

  response.write "                              </p>" & vbcrlf
  response.write "                          </td>" & vbcrlf
  response.write "                      </tr>" & vbcrlf
  response.write "                    </table>" & vbcrlf
  response.write "                </td>" & vbcrlf
  response.write "            </tr>" & vbcrlf
 'END: Footer -----------------------------------------------------------------

  response.write "          </table>" & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

  displayButtons "PREVIEW"

  response.write "</p>" & vbcrlf
  response.write "</fieldset>" & vbcrlf
 '-----------------------------------------------------------------------------
 'END: CommunityLink Preview
 '-----------------------------------------------------------------------------
%>
<!--#Include file="../admin_footer.asp"--> 
<%
  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts & vbcrlf
     response.write "</script>" & vbcrlf
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
sub displayCommunityLinkFeatureOptions(p_rowType, p_rowcount, p_totalrows, p_bgcolor, p_scripts, p_featureid, _
                                       p_featurename, p_featurename_original, p_portalcolumn, p_displayorder, p_numListItemsShown, _
                                       p_numListItemsShown_original, p_showSection, p_sectionheader_bgcolor, _
                                       p_sectionheader_linecolor, p_sectionheader_fonttype, p_sectionheader_fontcolor, _
                                       p_sectiontext_bgcolor, p_sectiontext_fonttype, p_sectiontext_fontcolor, p_sectionlinks_align, _
                                       p_sectionlinks_fonttype, p_sectionlinks_fontcolor, p_sectionlinks_fontcolorhover)

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
  lcl_scripts = p_scripts & "enableDisableOptions('" & p_rowType & "'," & p_rowcount & ");" & vbcrlf

 'Determine if this feature is "turned-on" to be displayed on this org's Community Link screen.
  if p_showSection then
     lcl_checked_showSection = " checked=""checked"""
  else
     lcl_checked_showSection = ""
  end if

  if lcl_rowType = "CL" then
     response.write "                  <td style=""padding-left:10px;" & lcl_row_border & """>" & vbcrlf
     response.write "                      <input type=""hidden"" name=""featureid_"                  & p_rowcount & """ id=""featureid_"                  & p_rowcount & """ value=""" & p_featureid                  & """ size=""5"" maxlength=""10"" />" & vbcrlf
     response.write "                      <input type=""hidden"" name=""featurename_original_"       & p_rowcount & """ id=""featurename_original_"       & p_rowcount & """ value=""" & p_featurename_original       & """ size=""5"" maxlength=""255"" />" & vbcrlf
     response.write "                      <input type=""hidden"" name=""numListItemsShown_original_" & p_rowcount & """ id=""numListItemsShown_original_" & p_rowcount & """ value=""" & p_numListItemsShown_original & """ size=""3"" maxlength=""10"" />" & vbcrlf

    'Feature Name
     response.write "                      <input type=""text"" name=""featurename_" & p_rowcount & """ id=""featurename_" & p_rowcount & """ value=""" & p_featurename & """ size=""40"" maxlength=""255"" /><br />" & vbcrlf
     response.write "                      <table border=""0"" cellspacing=""0"" cellpadding=""2"" align=""right"" style=""background-color:" & p_bgcolor & "; margin-left:10px; margin-top:5px;"">" & vbcrlf

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

    'Reset Button
     response.write "                        <tr id=""styleProperties_ResetButton_" & p_rowcount & """>" & vbcrlf
     response.write "                            <td>Reset to Defaults:&nbsp;" & vbcrlf
     response.write "                            <td><input type=""button"" name=""resetButton_" & p_rowcount & """ id=""resetButton_" & p_rowcount & """ value=""Reset"" class=""button"" onclick=""resetFields('" & p_rowcount & "');"" /></td>" & vbcrlf
     response.write "                        </tr>" & vbcrlf

     response.write "                      </table>" & vbcrlf
     response.write "                  </td>" & vbcrlf
  end if

  if lcl_rowType = "SAVVY" then
     lcl_td_style         = "border-left:1pt solid #000000;"
     lcl_showSectionLabel = "Style"
  else
     lcl_td_style         = ""
     lcl_showSectionLabel = "Display"
  end if

  response.write "                  <td style=""color:#800000;padding-left:10px;" & lcl_td_style & lcl_row_border & """ nowrap=""nowrap"">" & vbcrlf

 'Show Section
  response.write "                      <div>" & lcl_showSectionLabel & ":&nbsp;" & vbcrlf
  response.write "                        <input type=""checkbox"" name=""showSection_" & lcl_rowType & "_" & p_rowcount & """ id=""showSection_" & lcl_rowType & "_" & p_rowcount & """ value=""on"" onclick=""enableDisableOptions('" & p_rowType & "','" & p_rowcount & "');""" & lcl_checked_showSection & " />" & vbcrlf
  response.write "                      </div>" & vbcrlf
  response.write "                      <table id=""styleProperties_" & lcl_rowType & "_" & p_rowcount & """ border=""0"" cellspacing=""0"" cellpadding=""2"" style=""background-color:" & p_bgcolor & "; margin-top:2px; margin-left:4px;"">" & vbcrlf

 'Number of List Items Shown
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder""># List Items:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder""><input type=""text"" name=""numListItemsShown_" & lcl_rowType & "_" & p_rowcount & """ id=""numListItemsShown_" & lcl_rowType & "_" & p_rowcount & """ value=""" & p_numListItemsShown & """ size=""3"" maxlength=""10"" onchange=""clearMsg('numListItemsShown_" & lcl_rowType & "_" & p_rowcount & "');"" /></td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'BEGIN: Header Options -------------------------------------------------------
 'Section Header - Background Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Header BG Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionheader_bgcolor_" & lcl_rowType & "_" & p_rowcount, p_sectionheader_bgcolor, 1
                                                  lcl_scripts = p_scripts & "changePreviewColor('sectionheader_bgcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Section Header - Line Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Line BG Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionheader_linecolor_" & lcl_rowType & "_" & p_rowcount, p_sectionheader_linecolor, 1
                                                  lcl_scripts = p_scripts & "changePreviewColor('sectionheader_linecolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
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
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">Header Font Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">" & vbcrlf
                                                  setupColorSelection "sectionheader_fontcolor_" & lcl_rowType & "_" & p_rowcount, p_sectionheader_fontcolor, 1
                                                  lcl_scripts = p_scripts & "changePreviewColor('sectionheader_fontcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
 'END: Header Options ---------------------------------------------------------

 'BEGIN: Section Text Options -------------------------------------------------
 'Section Text - Background Color
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Text BG Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectiontext_bgcolor_" & lcl_rowType & "_" & p_rowcount, p_sectiontext_bgcolor, 1
                                                  lcl_scripts = p_scripts & "changePreviewColor('sectiontext_bgcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
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
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">Text Font Color:</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"" class=""communityLink_bottomborder"">" & vbcrlf
                                                  setupColorSelection "sectiontext_fontcolor_" & lcl_rowType & "_" & p_rowcount, p_sectiontext_fontcolor, 1
                                                  lcl_scripts = p_scripts & "changePreviewColor('sectiontext_fontcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
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
                                                  lcl_scripts = p_scripts & "changePreviewColor('sectionlinks_fontcolor_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf

 'Link Row - Font Color Hover
  response.write "                        <tr>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">Link Row - Font Color:<br />(mouseover)</td>" & vbcrlf
  response.write "                            <td nowrap=""nowrap"">" & vbcrlf
                                                  setupColorSelection "sectionlinks_fontcolorhover_" & lcl_rowType & "_" & p_rowcount, p_sectionlinks_fontcolorhover, 1
                                                  lcl_scripts = p_scripts & "changePreviewColor('sectionlinks_fontcolorhover_" & lcl_rowType & "_" & p_rowcount & "');" & vbcrlf
  response.write "                            </td>" & vbcrlf
  response.write "                        </tr>" & vbcrlf
 'END: Link Row Options -------------------------------------------------------

  response.write "                      </table>" & vbcrlf
  response.write "                  </td>" & vbcrlf

end sub
%>
