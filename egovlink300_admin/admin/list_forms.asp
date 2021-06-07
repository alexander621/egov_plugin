<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("action line") = "Y" then
    response.redirect "outage_feature_offline.asp"
 end if

 sLevel = "../"  'Override of value from common.asp

 if not UserHasPermission(session("userid"),"form creator") then
	   response.redirect sLevel & "permissiondenied.asp"
 end if

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for search options
 lcl_scFormName                 = ""
 lcl_scFormType                 = ""
 lcl_scFormStatus               = ""
 lcl_scFormDisplayOnList        = ""
 lcl_scFormID                   = ""
 sSelectedFormType_internal     = ""
 sSelectedFormType_public       = ""
 sSelectedFormStatus_on         = ""
 sSelectedFormStatus_off        = ""
 sSelectedFormDisplayOnList_on  = ""
 sSelectedFormDisplayOnList_off = ""

 if trim(request("scFormName")) <> "" then
    lcl_scFormName = trim(request("scFormName"))
 end if

 if trim(request("scFormType")) <> "" then
    lcl_scFormType = trim(request("scFormType"))
    lcl_scFormType = ucase(lcl_scFormType)

    if lcl_scFormType = "INTERNAL ONLY" then
       sSelectedFormType_internal = " selected=""selected"""
    elseif lcl_scFormType = "PUBLIC" then
       sSelectedFormType_public = " selected=""selected"""
    end if
 end if

 if trim(request("scFormStatus")) <> "" then
    lcl_scFormStatus = trim(request("scFormStatus"))
    lcl_scFormStatus = ucase(lcl_scFormStatus)

    if lcl_scFormStatus = "ON" then
       sSelectedFormStatus_on = " selected=""selected"""
    elseif lcl_scFormStatus = "OFF" then
       sSelectedFormStatus_off = " selected=""selected"""
    end if
 end if

 if trim(request("scDisplayOnList")) <> "" then
    lcl_scFormDisplayOnList = trim(request("scDisplayOnList"))
    lcl_scFormDisplayOnList = ucase(lcl_scFormDisplayOnList)

    if lcl_scFormDisplayOnList = "ON" then
       sSelectedFormDisplayOnList_on = " selected=""selected"""
    elseif lcl_scFormDisplayOnList = "OFF" then
       sSelectedFormDisplayOnList_off = " selected=""selected"""
    end if
 end if

 if trim(request("scFormID")) <> "" then
    lcl_scFormID = trim(request("scFormID"))
 end if
%>
<html>
<head>
 	<title>E-Gov Administration Console {Forms Management}</title>

	 <link type="text/css" rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	 <link type="text/css" rel="stylesheet" href="../global.css">

<style type="text/css">
.fieldset
{
   margin: 10px 0px;
   border-radius: 6px;
}

.fieldset legend
{
   padding: 4px 8px;
   border: 1pt solid #808080;
   border-radius: 6px;
   font-size: 1.25em;
   color: #800000;
}

#formsCreatorPageHeader
{
   font-size: 1.25em;
   font-weight: bold;
}

#screenMsg
{
   text-align: right;
   font-size: 1.125em;
   font-weight: bold;
   color: #ff0000;
}

#scFormName
{
   width: 83%;
}

#buttonSearch,
#buttonCreate
{
   cursor: pointer;
}

#buttonCreate
{
   margin: 10px 0px;
}

.formLabel
{
   white-space: nowrap;
}

#noFormsExist
{
   margin: 10px 0px;
   text-align: center;
   font-size: 1.25em;
   font-weight: bold;
   color: #ff0000;
}

#formsListTable
{
   width: 100%;
   background-color: #eeeeee;
}

#formsListTable th
{
   padding: 2px;
}

</style>

  <script src="../scripts/ajaxLib.js"></script>
  <script src="../scripts/modules.js"></script>
  <script src="../scripts/jquery-1.9.1.min.js"></script>

<script>
<!--
$(document).ready(function() {

  $('#scFormName').focus();

  $('#buttonSearch').click(function() {
     $('#searchForms').submit();
  });

  $('#buttonCreate').click(function() {
     var lcl_url  = 'copy_form.asp';
         lcl_url += '?task=NEW';
         lcl_url += '&iformid=57';
         lcl_url += '&iorgid=<%=session("orgid")%>';

     location.href = lcl_url;
  });

});


function editForm(iFormID)
{
   var lcl_url  = 'manage_form.asp';
       lcl_url += '?iformid=' + iFormID;

   location.href = lcl_url;
}

function changeFormOption(iFormID, iTask)
{
   var sTask = '';

   if(iTask != '')
   {
      sTask = iTask.toUpperCase();
   }

//alert('maintainFormOptions.asp?iformid=' + iFormID + '&task=INTERNAL');
   $.post('maintainFormOptions.asp', {
      formid: iFormID,
      task:    sTask
   }, function(result) {
      if(result == 'SU')
      {
         var sFieldID      = '';
         var sCurrentValue = '';
         var sCurrentColor = '';

         if(sTask == 'INTERNAL')
         {
            sFieldID = 'formTypeLabel_' + iFormID;
         }
         else if(sTask == 'ENABLEFORM')
         {
            sFieldID = 'formStatusLabel_' + iFormID;
         }
         else if(sTask == 'DISPLAYONLIST')
         {
            sFieldID = 'formDisplayOnListLabel_' + iFormID;
         }

         sCurrentValue = $('#' + sFieldID).html();

         if(sCurrentValue == 'Public')
         {
            sCurrentValue = 'Internal Only';
            sCurrentColor = '#ff0000';
         }
         else if(sCurrentValue == 'Internal Only')
         {
            sCurrentValue = 'Public';
            sCurrentColor = '#0000ff';
         }
         else if(sCurrentValue == 'ON')
         {
            sCurrentValue = 'OFF';
            sCurrentColor = '#ff0000';
         }
         else if(sCurrentValue == 'OFF')
         {
            sCurrentValue = 'ON';
            sCurrentColor = '#008000';
         }
         else if(sCurrentValue == 'On List')
         {
            sCurrentValue = 'Not on List';
            sCurrentColor = '#ff0000';
         }
         else if(sCurrentValue == 'Not on List')
         {
            sCurrentValue = 'On List';
            sCurrentColor = '#0000ff';
         }

         $('#' + sFieldID).html(sCurrentValue);
         $('#' + sFieldID).css('color', sCurrentColor);
      }
   });
}

 	function confirm_delete(iActionFormID) {
    sFormName      = document.getElementById("formname_" + iActionFormID).innerHTML;
    //sDeleteMessage = "Are you sure you want to delete form: (" + iActionFormID + ") - " + sFormName + "?\nAll related data will be lost.";
    sDeleteMessage = "Are you sure you want to delete form: (" + iActionFormID + ") - " + sFormName + "?";

  		if (confirm(sDeleteMessage))	{ 
    				//DELETE HAS BEEN VERIFIED
    				location.href='delete_form.asp?iformid=' + iActionFormID;
 			}
 	}

function changeRowColor(pID,pStatus) {
  if(pStatus=="OVER") {
     document.getElementById(pID).style.cursor          = "hand";
     document.getElementById(pID).style.backgroundColor = "#93bee1";
  }else{
     document.getElementById(pID).style.cursor          = "";
     document.getElementById(pID).style.backgroundColor = "";
  }
}

function updateShowInActionLineSearch(iFormID) {
  if(document.getElementById("showInALSearch" + iFormID).checked == true) {
     lcl_search_value = "on";
  }else{
     lcl_search_value = "off";
  }

  //Build the parameter string
		var sParameter  = 'isAjaxRoutine=Y';
  sParameter     += '&formid='   + encodeURIComponent(iFormID);
  sParameter     += '&ALSearch=' + encodeURIComponent(lcl_search_value);
  doAjax('updateShowInALSearch.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "&nbsp;";
}
//-->
</script>
</head>
<body onload="<%=lcl_onload%>">
	<% ShowHeader sLevel %>
	<!-- #include file="../menu/menu.asp" //-->
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf

  response.write "<div id=""formsCreatorPageHeader"">E-Gov Action Line Forms Creator</div>" & vbcrlf
  response.write "<div id=""screenMsg"">&nbsp;</div>" & vbcrlf

 'BEGIN: Search Criteria ------------------------------------------------------
  response.write "<fieldset class=""fieldset"">" & vbcrlf
  response.write "  <legend>Search Options</legend>" & vbcrlf
  response.write "  <form name=""searchForms"" id=""searchForms"" method=""post"" action=""list_forms.asp"">" & vbcrlf
  response.write "  <table border=""0"">" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td class=""formLabel"">Form Name:</td>" & vbcrlf
  response.write "        <td width=""100%"" colspan=""3"">" & vbcrlf
  response.write "            <input type=""text"" name=""scFormName"" id=""scFormName"" value=""" & lcl_scFormName & """ />" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td>Form Type:</td>" & vbcrlf
  response.write "        <td>" & vbcrlf
  response.write "            <select name=""scFormType"" id=""scFormType"">" & vbcrlf
  response.write "              <option value=""""></option>" & vbcrlf
  response.write "              <option value=""Internal Only""" & sSelectedFormType_internal & ">Internal Only</option>" & vbcrlf
  response.write "              <option value=""Public"""        & sSelectedFormType_public   & ">Public</option>" & vbcrlf
  response.write "            </select>" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "        <td>" & vbcrlf
  response.write "            Form Status:" & vbcrlf
  response.write "            <select name=""scFormStatus"" id=""scFormStatus"">" & vbcrlf
  response.write "              <option value=""""></option>" & vbcrlf
  response.write "              <option value=""ON"""  & sSelectedFormStatus_on  & ">On</option>" & vbcrlf
  response.write "              <option value=""OFF""" & sSelectedFormStatus_off & ">Off</option>" & vbcrlf
  response.write "            </select>" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "        <td>" & vbcrlf
  response.write "            Display in List:&nbsp;&nbsp;" & vbcrlf
  response.write "            <select name=""scDisplayOnList"" id=""scDisplayOnList"">" & vbcrlf
  response.write "              <option value=""""></option>" & vbcrlf
  response.write "              <option value=""ON"""  & sSelectedFormDisplayOnList_on  & ">On List</option>" & vbcrlf
  response.write "              <option value=""OFF""" & sSelectedFormDisplayOnList_off & ">Not on List</option>" & vbcrlf
  response.write "            </select>" & vbcrlf
  response.write "            <br />(public-side only)" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "    <tr>" & vbcrlf
  response.write "        <td>Form ID:</td>" & vbcrlf
  response.write "        <td>" & vbcrlf
  response.write "            <input type=""text"" name=""scFormID"" id=""scFormID"" value=""" & lcl_scFormID & """ size=""5"" />" & vbcrlf
  response.write "        </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "  </table>" & vbcrlf
  response.write "  </form>" & vbcrlf
  response.write "  <input type=""button"" name=""buttonSearch"" id=""buttonSearch"" value=""Search"" />" & vbcrlf
  response.write "</fieldset>" & vbcrlf
 'END: Search Criteria --------------------------------------------------------

 'BEGIN: Forms List -----------------------------------------------------------
  response.write "<div>" & vbcrlf
  response.write "  <input type=""button"" name=""create"" id=""buttonCreate"" value=""Create a New Form"" />" & vbcrlf
  response.write "  <button type=""button"" name=""market"" id=""buttonMarket"" style=""float:right;"" onclick=""window.location='formmarket.asp'"">View Forms Market</button>" & vbcrlf

                    subListForms session("orgid"), _
                                 lcl_scFormName, _
                                 lcl_scFormType, _
                                 lcl_scFormStatus, _
                                 lcl_scFormDisplayOnList, _
                                 lcl_scFormID

  response.write "</div>" & vbcrlf
 'END: Forms List -------------------------------------------------------------

  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf
%>
<!--#Include file="../admin_footer.asp"-->
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
sub subListForms(iOrgID, _
                 iSCFormName, _
                 iSCFormType, _
                 iSCFormStatus, _
                 iSCFormDisplayOnList, _
                 iSCFormID)

  dim sOrgID, lcl_orghasfeature_requestmergeforms
  dim sSCFormName, sSCFormType, sSCFormStatus, sSCFormDisplayOnList, sSCFormID

  lcl_orghasfeature_requestmergeforms = orghasfeature("requestmergeforms")
  sOrgID               = 0
  sSCFormName          = ""
  sSCFormType          = ""
  sSCFormStatus        = ""
  sSCFormDisplayOnList = ""
  sSCFormID            = ""
  sDBIsInternal        = ""
  sDBIsEnabled         = ""

 	if iOrgID <> "" then
     if not containsApostrophe(iOrgID) then
        sOrgID = clng(iOrgID)
     end if
 	end if

  if iSCFormName <> "" then
     sSCFormName = ucase(iSCFormName)
     sSCFormName = dbsafe(sSCFormName)
     sSCFormName = "'%" & sSCFormName & "%'"
  end if

  if iSCFormType <> "" then
     sSCFormType = ucase(iSCFormType)

     if sSCFormType = "INTERNAL ONLY" then
        sDBIsInternal = "1"
     elseif sSCFormType = "PUBLIC" then
        sDBIsInternal = "0"
     end if
  end if

  if iSCFormStatus <> "" then
     sSCFormStatus = ucase(iSCFormStatus)

     if sSCFormStatus = "ON" then
        sDBIsEnabled = "1"
     elseif sSCFormStatus = "OFF" then
        sDBIsEnabled = "0"
     end if
  end if

  if iSCFormDisplayOnList <> "" then
     sSCFormDisplayOnList = ucase(iSCFormDisplayOnList)

     if sSCFormDisplayOnList = "ON" then
        sDBIsOnList = "1"
     elseif sSCFormDisplayOnList = "OFF" then
        sDBIsOnList = "0"
     end if
  end if

 	if iSCFormID <> "" then
     if not containsApostrophe(iSCFormID) then
        sSCFormID = clng(iSCFormID)
        sSCFormID = "'%" & sSCFormID & "%'"
     end if
 	end if

 	sSQL = "SELECT f.*, "
  sSQL = sSQL & " (select count(r.action_autoid) "
  sSQL = sSQL &  " from egov_actionline_requests r "
  sSQL = sSQL &  " where r.category_id = f.action_form_id) as totalrequests "
  sSQL = sSQL & " FROM egov_action_request_forms f "
  sSQL = sSQL & " WHERE (f.action_form_type <> 2) "
  sSQL = sSQL & " AND f.orgid = " & sOrgID

  if sSCFormName <> "" then
     sSQL = sSQL & " AND upper(f.action_form_name) like (" & sSCFormName & ") "
  end if

  if sDBIsInternal <> "" then
     sSQL = sSQL & " AND isnull(f.action_form_internal, 0) = " & sDBIsInternal
  end if

  if sDBIsEnabled <> "" then
     sSQL = sSQL & " AND isnull(f.action_form_enabled, 0) = " & sDBIsEnabled
  end if

  if sDBIsOnList <> "" then
     sSQL = sSQL & " AND isnull(f.action_form_displayOnList, 0) = " & sDBIsOnList
  end if

  if sSCFormID <> "" then
     sSQL = sSQL & " AND f.action_form_id like (" & sSCFormID & ") "
  end if

  sSQL = sSQL & " ORDER BY f.action_form_type, f.action_form_name "

 	set oFormList = Server.CreateObject("ADODB.Recordset")
 	oFormList.Open sSQL, Application("DSN"), 3, 1
	
 	if not oFormList.eof then
    	response.write "<table id=""formsListTable"" border=""0"" cellspacing=""0"" class=""tablelist"">" & vbcrlf
	 	  response.write "  <tr>" & vbcrlf
     response.write "      <th align=""left"">Form Name</th>" & vbcrlf
     response.write "      <th>Form<br />ID</th>" & vbcrlf
     response.write "      <th>Form<br />Type</th>" & vbcrlf
     response.write "      <th>Form<br />Status</th>" & vbcrlf
     response.write "      <th>Display<br />in List<br />(public-side)</th>" & vbcrlf

     if lcl_orghasfeature_requestmergeforms then
        response.write "      <th>Public PDF</th>" & vbcrlf
     end if

     response.write "      <th align=""center"" nowrap=""nowrap"">Include in<br />Action Line<br />Search</th>" & vbcrlf
     response.write "      <th>&nbsp;</th>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     sRowClass = "formrowW"

   		do while not oFormList.eof
        sRowClass           = changeBGColor(sRowClass,"formrowW","formrowG")
     			sType               = "STANDARD"
        sEnabledColor       = "#ff0000"
        sEnabledLabel       = "OFF"
        sDisplayOnListColor = "#ff0000"
        sDisplayOnListLabel = "Not on List"
        'blnEnabled          = 1
        lcl_showIncludeSearch_checkbox = True

     		'Determine the form type
      		if oFormList("action_form_type") = "1" then
       				sType = "CUSTOM"
        end if

     		'Determine if the form is enabled
      		if oFormList("action_form_enabled") then
           sEnabledColor = "#008000"
           sEnabledLabel = "ON"
           'blnEnabled    = 0
           lcl_showIncludeSearch_checkbox = False
        end if

     		'Determine if the form is displayed on the list (public-side)
      		if oFormList("action_form_displayOnList") then
           sDisplayOnListColor = "#0000ff"
           sDisplayOnListLabel = "On List"
        end if

    				'sEnabled  = "<font style=""color:" & sEnabledColor  & ";font-size:10px;"">" & sEnabledLabel  & "</font>"

       'Setup the javascript events for the row.
        lcl_row_onmouseover           = " onMouseOver=""changeRowColor('row_" & oFormList("action_form_id") & "','OVER')"" style=""cursor:pointer;"""
        lcl_row_onmouseout            = " onMouseOut=""changeRowColor('row_" & oFormList("action_form_id") & "','OUT')"""
        'lcl_row_onclick               = " onClick=""location.href='manage_form.asp?iformid=" & oFormList("action_form_id") & "';"""
        lcl_row_onclick               = " onClick=""editForm(" & oFormList("action_form_id") & ");"""
        lcl_row_onclick_internal      = " onclick=""changeFormOption(" & oFormList("action_form_id") & ", 'INTERNAL');"""
        lcl_row_onclick_enabled       = " onclick=""changeFormOption(" & oFormList("action_form_id") & ", 'ENABLEFORM');"""
        lcl_row_onclick_displayOnList = " onclick=""changeFormOption(" & oFormList("action_form_id") & ", 'DISPLAYONLIST');"""

       'Determine if org allows internal/public selection
        sDisplayOptionFormType       = "&nbsp;"
        'sDisplayOptionEnabled        = "<a href=""edit_form.asp?iorgid=" & oFormList("action_form_id") & "&iformid="&oFormList("action_form_id")&"&task=ENABLEFORM&blnenabled=" & blnEnabled & """>" & sEnabled & "</a>"
        sDisplayOptionEnabled        = "<a id=""formStatusLabel_"        & oFormList("action_form_id") & """ style=""color:" & sEnabledColor       & ";"">" & sEnabledLabel       & "</a>"
        sDisplayOptionDisplayOnList  = "<a id=""formDisplayOnListLabel_" & oFormList("action_form_id") & """ style=""color:" & sDisplayOnListColor & ";"">" & sDisplayOnListLabel & "</a>"
        sDisplayOptionShowInAL       = "&nbsp;<input type=""hidden"" name=""showInALSearch" & oFormList("action_form_id") & """ id=""showInALSearch" & oFormList("action_form_id") & """ value=""off"" />"
        sDisplayDeleteButton         = "&nbsp;"

       'Form Type: Internal Only / Public
        if session("OrgInternalEntry") then
           sInternalColor = "#0000ff"
           sInternalLabel = "Public"
           blnInternal    = 1

        		'Determine if the form is "public" or "internal only"
         		if oFormList("action_form_internal") then
              sInternalColor = "#ff0000"
              sInternalLabel = "Internal Only"
          				blnInternal    = 0
           end if 

       				'sInternal              = "<span style=""color:" & sInternalColor & ";"">" & sInternalLabel & "</span>"
    		     'sDisplayOptionFormType = "<a href=""edit_form.asp?iorgid=" & oFormList("action_form_id") & "&iformid="&oFormList("action_form_id")&"&task=INTERNAL&blnInternal=" & blnInternal & """>" & sInternal & "</a>"
           sDisplayOptionFormType = "<a id=""formTypeLabel_" & oFormList("action_form_id") & """ style=""color:" & sInternalColor & ";"">" & sInternalLabel & "</span>"
        end if

        if lcl_showIncludeSearch_checkbox then
           lcl_checked_showInALSearch = ""

           if oFormList("showInALSearch") then
              lcl_checked_showInALSearch = " checked=""checked"""
           end if

           sDisplayOptionShowInAL = "<input type=""checkbox"" name=""showInALSearch" & oFormList("action_form_id") & """ id=""showInALSearch" & oFormList("action_form_id") & """ value=""on"" onclick=""updateShowInActionLineSearch('" & oFormList("action_form_id") & "')""" & lcl_checked_showInALSearch & " />"
        end if

        if oFormList("totalrequests") < 1 then
           sDisplayDeleteButton = "<input type=""button"" name=""deleteButton"" id=""deleteButton"" value=""Delete"" class=""button"" onclick=""confirm_delete('" & oFormList("action_form_id") & "');"" />"
        end if

       'Display the row
   	  		response.write "  <tr class=""" & sRowClass & """ id=""row_" & oFormList("action_form_id") & """>" & vbcrlf
   	  		response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & lcl_row_onclick & " id=""formname_" & oFormList("action_form_id") & """>&nbsp;&nbsp;" & oFormList("action_form_name")
			if oFormList("formobile") then
				response.write "<br />&nbsp;&nbsp;<i>App Name:" & oFormList("mobilename") & "</i>"
			end if
			response.write "</td>" & vbcrlf
 		     response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & " align=""center"">(" & oFormList("action_form_id") & ")&nbsp;</td>" & vbcrlf
 		     response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & lcl_row_onclick_internal      & " nowrap=""nowrap"" align=""center"">" & sDisplayOptionFormType      & "</td>" & vbcrlf
 		     response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & lcl_row_onclick_enabled       & " nowrap=""nowrap"" align=""center"">" & sDisplayOptionEnabled       & "</td>" & vbcrlf
 		     response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & lcl_row_onclick_displayOnList & " nowrap=""nowrap"" align=""center"">" & sDisplayOptionDisplayOnList & "</td>" & vbcrlf

        if lcl_orghasfeature_requestmergeforms then
           sDisplayPublicPDF = "&nbsp;"

           if trim(RIGHT(oFormList("public_actionline_pdf"),30)) <> "" then
              sDisplayPublicPDF = trim(RIGHT(oFormList("public_actionline_pdf"),30))
           end if

           response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & ">" & sDisplayPublicPDF & "</td>" & vbcrlf
        end if

        response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & " align=""center"">"  & sDisplayOptionShowInAL & "</td>" & vbcrlf
 		     response.write "      <td class=""formlist""" & lcl_row_onmouseover & lcl_row_onmouseout & " nowrap=""nowrap"">" & sDisplayDeleteButton   & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

      		oFormList.movenext
     loop

    	response.write "</table>" & vbcrlf
 	else
     response.write "<div id=""noFormsExist"">*** No forms exist ***</div>" & vbcrlf
  end if

 	set oFormList = nothing

end sub

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "RSS_SUCCESS" then
        lcl_return = "Successfully Sent to RSS..."
     elseif iSuccess = "RSS_ERROR" then
        lcl_return = "ERROR: Failed to send to RSS..."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function
%>
