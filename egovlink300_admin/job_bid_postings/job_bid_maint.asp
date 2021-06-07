<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="job_bid_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME:  job_bid_maint.asp
' AUTHOR:    David Boyer
' CREATED:   01/29/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  01/29/08  David Boyer - INITIAL VERSION
' 1.1  05/20/09  David Boyer - Modified "Download Available" "Find a Link" button to include user-click tracking.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("job_postings,bid_postings") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

'INITIALIZE VARIABLES
 Dim sName, sDescription,blnDisplay
 Dim iDLid

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=hide

'Retrieve the search parameters
 lcl_sc_title             = request("sc_title")
 lcl_sc_status_id         = request("sc_status_id")
 lcl_sc_publicly_viewable = request("sc_publicly_viewable")
 lcl_sc_list_type         = request("sc_list_type")
 lcl_sc_orderby           = request("sc_orderby")
 lcl_autosend_email       = request("autosend_email")

 if lcl_autosend_email = "" then
    lcl_autosend_email = "N"
 end if

'Determine if user is adding/editing a job/bid posting
if request("posting_id") = "" OR NOT isnumeric(request("posting_id")) OR request("posting_id") = 0 then
	 '-- ADD --------------------------------
	  lcl_posting_id        = 0
   lcl_screen_mode_label = "Add"

   if checkForPermission(lcl_posting_id,lcl_sc_list_type) = "N" then
      response.redirect sLevel & "permissiondenied.asp"
   end if

   if lcl_sc_list_type = "JOB" then
      sTitle              = "Add New Job Posting"
      lcl_label           = "Job"
   elseif lcl_sc_list_type = "BID" then
      sTitle              = "Add New Bid Posting"
      lcl_label           = "Bid"
   end if

	  sLinkText = "Add"
else
	 '-- EDIT -------------------------------
	  lcl_posting_id        = request("posting_id")
   lcl_screen_mode_label = "Save"

   if checkForPermission(lcl_posting_id,lcl_sc_list_type) = "N" then
      response.redirect sLevel & "permissiondenied.asp"
   end if

   if lcl_sc_list_type = "JOB" then
      sTitle    = "Edit Job Posting"
      lcl_label = "Job"
   elseif lcl_sc_list_type = "BID" then
      sTitle    = "Edit Bid Posting"
      lcl_label = "Bid"
   end if

	  sLinkText = "Save"
end if

'Retrieve data for this posting.
	sSQL = "SELECT jb.jobbid_id, jb.posting_id, jb.posting_type, jb.title, jb.start_date, jb.end_date, jb.status_id, s.status_name, "
 sSQL = sSQL & " jb.additional_status_info, jb.description, jb.qualifications, jb.special_requirements, jb.misc_info, jb.active_flag, "

 if lcl_sc_list_type = "JOB" then
    sSQL = sSQL & " jb.job_salary, public_apply_for_position_actionline, "
 elseif lcl_sc_list_type = "BID" then
    sSQL = sSQL & " jb.bid_publication_info, jb.bid_submittal_info, jb.bid_opening_info, jb.bid_recipient, jb.bid_addendum_date, "
    sSQL = sSQL & " jb.bid_pre_bid_meeting, jb.bid_contact_person, jb.bid_fee, jb.bid_plan_spec_available, "
    sSQL = sSQL & " jb.bid_business_hours, jb.bid_fax_number, jb.bid_plan_holders, "
 end if

 sSQL = sSQL & " jb.download_available "
 sSQL = sSQL & " FROM egov_jobs_bids jb, egov_statuses s"
 sSQL = sSQL & " WHERE jb.status_id = s.status_id "
 sSQL = sSQL & " AND s.status_type = '" & lcl_sc_list_type & "'"
 sSQL = sSQL & " AND jb.posting_id = "  & lcl_posting_id
 sSQL = sSQL & " AND jb.orgid = "       & session("orgid")

	set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN") , 3, 1

	If NOT oValues.EOF Then
    lcl_jobbid_id                  = oValues("jobbid_id")
  		lcl_posting_type               = oValues("posting_type")
		  lcl_title                      = oValues("title")

    if oValues("start_date") <> "" then
       if CDate(oValues("start_date")) = CDate("1/1/1900") then
          lcl_start_date = ""
       else
        		lcl_start_date = datevalue(oValues("start_date"))
       end if

       lcl_start_hour = hour(oValues("start_date"))
       lcl_start_min  = minute(oValues("start_date"))
       lcl_start_ampm = right(oValues("start_date"),2)

    else
       lcl_start_date = ""
       lcl_start_hour = ""
       lcl_start_min  = ""
       lcl_start_ampm = ""
    end if

    if oValues("end_date") <> "" then
       if CDate(oValues("end_date")) = CDate("1/1/1900") then
          lcl_end_date = ""
       else
        		lcl_end_date = datevalue(oValues("end_date"))
       end if

       lcl_end_hour = hour(oValues("end_date"))
       lcl_end_min  = minute(oValues("end_date"))
       lcl_end_ampm = right(oValues("end_date"),2)

    else
       lcl_end_date = ""
       lcl_end_hour = ""
       lcl_end_min  = ""
       lcl_end_ampm = ""
    end if

    lcl_status_id                  = oValues("status_id")
    lcl_status_name                = oValues("status_name")
    lcl_additional_status_info     = oValues("additional_status_info")
    lcl_description                = oValues("description")
    lcl_qualifications             = oValues("qualifications")
    lcl_special_requirements       = oValues("special_requirements")
    lcl_misc_info                  = oValues("misc_info")
    lcl_active_flag                = oValues("active_flag")
    lcl_download_available         = oValues("download_available")

   'Retrieve the columns related to the listtype
    if lcl_sc_list_type = "JOB" then
       lcl_job_salary              = oValues("job_salary")
       lcl_public_apply_for_position_actionline = oValues("public_apply_for_position_actionline")
    elseif lcl_sc_list_type = "BID" then
       lcl_bid_publication_info    = oValues("bid_publication_info")
       lcl_bid_submittal_info      = oValues("bid_submittal_info")
       lcl_bid_opening_info        = oValues("bid_opening_info")
       lcl_bid_recipient           = oValues("bid_recipient")
       'lcl_bid_addendum_date       = REPLACE(oValues("bid_addendum_date"),"1/1/1900","")
       lcl_bid_pre_bid_meeting     = oValues("bid_pre_bid_meeting")
       lcl_bid_contact_person      = oValues("bid_contact_person")
       lcl_bid_fee                 = oValues("bid_fee")
       lcl_bid_plan_spec_available = oValues("bid_plan_spec_available")
       lcl_bid_business_hours      = oValues("bid_business_hours")
       lcl_bid_fax_number          = oValues("bid_fax_number")
       lcl_bid_plan_holders        = oValues("bid_plan_holders")

       if oValues("bid_addendum_date") <> "" then
          if CDate(oValues("bid_addendum_date")) = CDate("1/1/1900") then
             lcl_bid_addendum_date = ""
          else
           		lcl_bid_addendum_date = oValues("bid_addendum_date")
          end if
       else
          lcl_bid_addendum_date = ""
       end if

    end if
 else
    lcl_jobbid_id                  = ""
  		lcl_posting_type               = lcl_sc_list_type
		  lcl_title                      = ""
  		lcl_start_date                 = ""
  		lcl_end_date                   = ""
    lcl_status_id                  = 0
    lcl_status_name                = ""
    lcl_additional_status_info     = ""
    lcl_description                = ""
    lcl_qualifications             = ""
    lcl_special_requirements       = ""
    lcl_misc_info                  = ""
    lcl_active_flag                = ""
    lcl_download_available         = ""

   'Retrieve the columns related to the listtype
    if lcl_sc_list_type = "JOB" then
       lcl_job_salary              = ""
       lcl_public_apply_for_position_actionline = ""
    elseif lcl_sc_list_type = "BID" then
       lcl_bid_publication_info    = ""
       lcl_bid_submittal_info      = ""
       lcl_bid_opening_info        = ""
       lcl_bid_recipient           = ""
       lcl_bid_addendum_date       = ""
       lcl_bid_pre_bid_meeting     = ""
       lcl_bid_contact_person      = ""
       lcl_bid_fee                 = ""
       lcl_bid_plan_spec_available = ""
       lcl_bid_business_hours      = ""
       lcl_bid_fax_number          = ""
       lcl_bid_plan_holders        = ""
    end if
 end if

	oValues.close
	Set oValues = nothing

'Check for a screen message
 lcl_success = request("success")
 lcl_onload  = ""

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Build the BODY onload
 lcl_onload = lcl_onload & "enableDisableTime('start_date');"
 lcl_onload = lcl_onload & "enableDisableTime('end_date');"
 lcl_onload = lcl_onload & "setMaxLength();"

'Check for org features
 lcl_orghasfeature_clickcounter_postings = orghasfeature("clickcounter_postings")
%>
<html>
<head>
	<title>E-Gov Administration Console {<%=lcl_label%> Postings}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script language="javascript" src="tablesort.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>
 <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

<script language="javascript">
 function doCalendar(sField) {
   var w = (screen.width - 350)/2;
   var h = (screen.height - 350)/2;
   eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=posting_maint", "_jobbid", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
 }

 //function doPicker(sFormField) {
 //  w = (screen.width - 350)/2;
 //  h = (screen.height - 350)/2;
 //  eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
 //}

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
     //lcl_showFolderStart = "&folderStart=published_documents";
     lcl_showFolderStart = "&folderStart=CITY_ROOT";
  }

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += "&returnAsHTMLLink=Y";

  <% if lcl_orghasfeature_clickcounter_postings then %>
  pickerURL += "&includeClickCounter=Y";
  <% end if %>

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
       textEl.value = textEl.value + text;
 }

function enableDisableTime(p_field) {
  if(document.getElementById(p_field).value=="") {
     document.getElementById(p_field+"_hour").disabled=true;
     document.getElementById(p_field+"_min").disabled=true;
     document.getElementById(p_field+"_ampm").disabled=true;
  }else{
     document.getElementById(p_field+"_hour").disabled=false;
     document.getElementById(p_field+"_min").disabled=false;
     document.getElementById(p_field+"_ampm").disabled=false;
  }
}

function validateFields() {
  var lcl_false_count = 0;
		var daterege        = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;

  lcl_startDate    = document.getElementById("start_date").value;
  lcl_endDate      = document.getElementById("end_date").value;

		var dateStartOk    = daterege.test(lcl_startDate);
		var dateEndOk      = daterege.test(lcl_endDate);

<% if lcl_sc_list_type = "BID" then %>
  lcl_addendumDate = document.getElementById("bid_addendum_date").value;
		var dateAddendumOk = daterege.test(lcl_addendumDate);

		if (lcl_addendumDate != "" && ! dateAddendumOk ) {
      document.getElementById("bid_addendum_date").focus();
      inlineMsg(document.getElementById("bid_addendum_date_pop").id,'<strong>Invalid Value: </strong> The "Addendum Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'bid_addendum_date_pop');
      lcl_false_count = lcl_false_count + 1
  }else{
      clearMsg("bid_addendum_date_pop");
  }
<% end if %>

		if (lcl_endDate != "" && ! dateEndOk ) {
      document.getElementById("end_date").focus();
      inlineMsg(document.getElementById("end_date_pop").id,'<strong>Invalid Value: </strong> The "End Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'end_date_pop');
      lcl_false_count = lcl_false_count + 1
  }else{
      clearMsg("end_date_pop");
  }

		if (lcl_startDate != "" && ! dateStartOk ) {
      document.getElementById("start_date").focus();
      inlineMsg(document.getElementById("start_date_pop").id,'<strong>Invalid Value: </strong> The "Start Date" must be in date format.<br /><span style="color:#800000;">(i.e. mm/dd/yyyy)</span>',10,'start_date_pop');
      lcl_false_count = lcl_false_count + 1
  }else{
      clearMsg("start_date_pop");
  }

  if(lcl_false_count > 0) {
     return false;
  }else{
     document.getElementById("posting_maint").submit();
     return true;
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
<body onLoad="<%=lcl_onload%>">
<%'DrawTabs tabRecreation,1%>

<% ShowHeader sLevel %>
<!--#include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<% ShowEmailWarning %>
	
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="start">
  <caption align="left" style="padding-bottom:10px;">
    <font size="+1"><strong><%=sTitle%></strong></font><br />
    <input type="button" name="backButton" id="backButton" value="Return to <%=lcl_label%> Postings" class="button" onclick="location.href='job_bid_list.asp?sc_title=<%=lcl_sc_title%>&sc_status_id=<%=lcl_sc_status_id%>&sc_publicly_viewable=<%=lcl_sc_publicly_viewable%>&sc_list_type=<%=lcl_sc_list_type%>&sc_orderby=<%=lcl_sc_orderby%>';" />
  </caption>
  <tr>
      <td>
          <% display_buttons lcl_sc_title, lcl_sc_status_id, lcl_sc_publicly_viewable, lcl_sc_list_type, lcl_sc_orderby, lcl_autosend_email %>
      </td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
</table>
<table border="0" cellpadding="5" cellspacing="0" class="tableadmin">
  <form name="posting_maint" id="posting_maint" action="job_bid_action.asp" method="post">
    <input type="<%=lcl_hidden%>" name="posting_id" value="<%=lcl_posting_id%>" />
    <input type="<%=lcl_hidden%>" name="posting_type" value="<%=lcl_posting_type%>" size="10" maxlength="50" />
    <input type="<%=lcl_hidden%>" name="sc_title" value="<%=lcl_sc_title%>" size="15" maxlength="512" />
    <input type="<%=lcl_hidden%>" name="sc_descripion" value="<%=lcl_sc_status_id%>" size="15" maxlength="1024" />
    <input type="<%=lcl_hidden%>" name="sc_publicly_viewable" value="<%=lcl_sc_publicly_viewable%>" size="15" maxlength="5" />
    <input type="<%=lcl_hidden%>" name="sc_list_type" value="<%=lcl_sc_list_type%>" size="15" maxlength="100" />
    <input type="<%=lcl_hidden%>" name="sc_orderby" value="<%=lcl_sc_orderby%>" size="15" maxlength="50" />
    <input type="<%=lcl_hidden%>" name="autosend_email" value="<%=lcl_autosend_email%>" size="1" maxlength="1" />
  <tr>
      <th align="left"><%=lcl_label%> Information</th>
<%
  if lcl_sc_list_type = "JOB" then
     response.write "<th align=""right"">" & vbcrlf
     response.write "    ""Apply for Position"" link - Action Line Request Form: " & vbcrlf
     response.write "    <select name=""public_apply_for_position_actionline"">" & vbcrlf
     response.write "      <option value=""""></option>" & vbcrlf

     sSQLf = "SELECT action_form_id, action_form_name "
     sSQLf = sSQLf & " FROM egov_action_request_forms "
     sSQLf = sSQLf & " WHERE orgid = " & session("orgid")
     sSQLf = sSQLf & " ORDER BY UPPER(action_form_name) "

     set rsf = Server.CreateObject("ADODB.Recordset")
     rsf.Open sSQLf, Application("DSN") , 3, 1

     if not rsf.eof then
        while not rsf.eof
           if lcl_public_apply_for_position_actionline <> "" then
              if CLng(lcl_public_apply_for_position_actionline) = CLng(rsf("action_form_id")) then
                 lcl_selected = " selected"
              else
                 lcl_selected = ""
              end if
           else
              lcl_selected = ""
           end if

           response.write "<option value=""" & rsf("action_form_id") & """" & lcl_selected & ">" & rsf("action_form_name") & "</option>" & vbcrlf
           rsf.movenext
        wend
     end if

     response.write "    </select>" & vbcrlf
     response.write "</th>" & vbcrlf
  else
     response.write "<input type=""" & lcl_hidden & """ name=""public_apply_for_position_actionline"" size=""3"" maxlength=""5"">" & vbcrlf
  end if
%>
  </tr>
  <tr>
      <td colspan="2">
          <table border="0" cellspacing="0" cellpadding="5">
            <tr>
                <td><%=lcl_label%> ID:</td>
                <td colspan="4"><input type="text" name="jobbid_id" value="<%=lcl_jobbid_id%>" size="50" maxlength="50"></td>
            </tr>
            <tr>
                <td>Title:</td>
                <td colspan="4"><input type="text" name="title" value="<%=lcl_title%>" size="50" maxlength="500"></td>
            </tr>
            <tr>
                <td nowrap="nowrap">Start Date:</td>
                <td nowrap="nowrap">
                    <input type="text" name="start_date" id="start_date" value="<%=lcl_start_date%>" size="10" maxlength="10" onblur="enableDisableTime('start_date');" onchange="clearMsg('start_date_pop');" />&nbsp;
                    <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" id="start_date_pop" height="16" width="16" border="0" onclick="javascript:void doCalendar('start_date');" /></span>
                    &nbsp;[mm/dd/yyyy]
                </td>
                <td nowrap="nowrap">
                    Start Time:
                    <select name="start_hour" id="start_date_hour" class="time">
                    <%
                      for hr = 1 to 12
                          if lcl_start_hour <> "" then
                             if lcl_start_hour > 12 then
                                lcl_start_hour = lcl_start_hour - 12
                             end if

                             if clng(lcl_start_hour) = clng(hr) then
                                lcl_selected_start = " selected"
                             else
                                lcl_selected_start = ""
                             end if
                          else
                             lcl_selected_start = ""
                          end if

                          response.write "<option value=""" & hr & """" & lcl_selected_start & ">" & hr & "</option>" & vbcrlf
                      next

                      response.write "</select>" & vbcrlf
                      response.write " : " & vbcrlf
                      response.write "<select name=""start_minute"" id=""start_date_min"" class=""time"">" & vbcrlf

                      min = 0
                      do while min < 56
                         if lcl_start_min <> "" then
                            if clng(lcl_start_min) = clng(min) then
                               lcl_selected_start = " selected"
                            else
                               lcl_selected_start = ""
                            end if
                         else
                            lcl_selected_start = ""
                         end if

                         if min < 10 then
                            min = "0" & min
                         end if

                         response.write "<option value=""" & min & """" & lcl_selected_start & ">" & min & "</option>" & vbcrlf

                         min = min + 5
                      loop
                     %>
                    </select>
                    <select name="start_ampm" id="start_date_ampm" class="time">
                    <%
                      if UCASE(lcl_start_ampm) = "PM" then
                         lcl_selected_am = ""
                         lcl_selected_pm = " selected"
                      else
                         lcl_selected_am = " selected"
                         lcl_selected_pm = ""
                      end if
                    %>
                      <option value="AM"<%=lcl_selected_am%>>AM</option>
                      <option value="PM"<%=lcl_selected_pm%>>PM</option>
                    </select>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td>End Date:</td>
                <td><input type="text" name="end_date" id="end_date" value="<%=lcl_end_date%>" size="10" maxlength="10" onblur="enableDisableTime('end_date');" onchange="clearMsg('end_date_pop');" />&nbsp;
                    <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" id="end_date_pop" height="16" width="16" border="0" onclick="javascript:void doCalendar('end_date');" /></span>
                    &nbsp;[mm/dd/yyyy]
                </td>
                <td nowrap="nowrap">
                    End Time:
                    <%
                      response.write "<select name=""end_hour"" id=""end_date_hour"" class=""time"">" & vbcrlf

                      for hr = 1 to 12
                          if lcl_end_hour <> "" then
                             if lcl_end_hour > 12 then
                                lcl_end_hour = lcl_end_hour - 12
                             end if

                             if clng(lcl_end_hour) = clng(hr) then
                                lcl_selected_end = " selected"
                             else
                                lcl_selected_end = ""
                             end if
                          else
                             lcl_selected_end = ""
                          end if

                          response.write "<option value=""" & hr & """" & lcl_selected_end & ">" & hr & "</option>" & vbcrlf
                      next

                      response.write "</select>" & vbcrlf
                      response.write " : " & vbcrlf
                      response.write "<select name=""end_minute"" id=""end_date_min"" class=""time"">" & vbcrlf

                      min = 0
                      do while min < 56
                         if lcl_end_min <> "" then
                            if clng(lcl_end_min) = clng(min) then
                               lcl_selected_end = " selected"
                            else
                               lcl_selected_end = ""
                            end if
                         else
                            lcl_selected_end = ""
                         end if

                         if min < 10 then
                            min = "0" & min
                         end if

                         response.write "<option value=""" & min & """" & lcl_selected_end & ">" & min & "</option>" & vbcrlf

                         min = min + 5
                      loop

                      response.write "</select>" & vbcrlf
                      response.write "<select name=""end_ampm"" id=""end_date_ampm"" class=""time"">" & vbcrlf

                      if UCASE(lcl_end_ampm) = "PM" then
                         lcl_selected_am = ""
                         lcl_selected_pm = " selected"
                      else
                         lcl_selected_am = " selected"
                         lcl_selected_pm = ""
                      end if

                      response.write "  <option value=""AM""" & lcl_selected_am & ">AM</option>" & vbcrlf
                      response.write "  <option value=""PM""" & lcl_selected_pm & ">PM</option>" & vbcrlf
                      response.write "</select>" & vbcrlf
                    %>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td>Status:</td>
                <td colspan="4">
                    <select name="status_id">
                      <% displayPostingStatuses session("orgid"), lcl_sc_list_type, lcl_status_id %>
                    </select>
                </td>
            </tr>
            <tr>
                <td>Additional Status Info:</td>
                <td colspan="4"><input type="text" name="additional_status_info" value="<%=lcl_additional_status_info%>" size="50" maxlength="500"></td>
            </tr>
            <tr>
                <td>Active:</td>
                <td colspan="4">
                    <%
                      if lcl_active_flag = "Y" then
                         lcl_selected_active_yes = " selected"
                         lcl_selected_active_no  = ""
                      elseif lcl_active_flag = "N" then
                         lcl_selected_active_yes = ""
                         lcl_selected_active_no  = " selected"
                      else
                         lcl_selected_active_yes = " selected"
                         lcl_selected_active_no  = ""
                      end if
                    %>
                    <select name="active_flag">
                      <option value="Y"<%=lcl_selected_active_yes%>>Yes</option>
                      <option value="N"<%=lcl_selected_active_no%>>No</option>
                    </select>
                </td>
            </tr>
         <%
           if lcl_sc_list_type = "BID" then
             'Publication Info
              response.write "<tr valign=""top"">" & vbcrlf
              response.write "    <td>Publication Info:</td>" & vbcrlf
              response.write "    <td colspan=""4""><input type=""text"" name=""bid_publication_info"" value=""" & lcl_bid_publication_info & """ size=""50"" maxlength=""500"" /></td>" & vbcrlf
              response.write "</tr>" & vbcrlf

             'Submittal Info
              response.write "<tr valign=""top"">" & vbcrlf
              response.write "    <td>Submittal Info:</td>" & vbcrlf
              response.write "    <td colspan=""4""><input type=""text"" name=""bid_submittal_info"" value=""" & lcl_bid_submittal_info & """ size=""50"" maxlength=""500"" /></td>" & vbcrlf
              response.write "</tr>" & vbcrlf

             'Bid Opening Info
              response.write "<tr valign=""top"">" & vbcrlf
              response.write "    <td>Bid Opening Info:</td>" & vbcrlf
              response.write "    <td colspan=""4""><input type=""text"" name=""bid_opening_info"" value=""" & lcl_bid_opening_info & """ size=""50"" maxlength=""500"" /></td>" & vbcrlf
              response.write "</tr>" & vbcrlf

             'Bid Recipient
              response.write "<tr valign=""top"">" & vbcrlf
              response.write "    <td>Bid Recipient:</td>" & vbcrlf
              response.write "    <td colspan=""4""><input type=""text"" name=""bid_recipient"" value=""" & lcl_bid_recipient & """ size=""50"" maxlength=""500"" /></td>" & vbcrlf
              response.write "    </tr>" & vbcrlf
           end if

          'Description
           response.write "<tr valign=""top"">" & vbcrlf
           response.write "    <td>Description:</td>" & vbcrlf
           response.write "    <td colspan=""4""><textarea name=""description"" rows=""5"" cols=""80"">" & lcl_description & "</textarea></td>" & vbcrlf
           response.write "</tr>" & vbcrlf

          'Salary
           if lcl_sc_list_type = "JOB" then
              response.write "<tr valign=""top"">" & vbcrlf
              response.write "    <td>Salary:</td>" & vbcrlf
              response.write "    <td colspan=""4""><textarea name=""job_salary"" rows=""5"" cols=""80"">" & lcl_job_salary & "</textarea></td>" & vbcrlf
              response.write "</tr>" & vbcrlf
           end if
         %>
            <tr valign="top">
                <td>Qualifications:</td>
                <td colspan="4"><textarea name="qualifications" rows="5" cols="80"><%=lcl_qualifications%></textarea></td>
            </tr>
            <tr valign="top">
                <td>Special Requirements:</td>
                <td colspan="4"><textarea name="special_requirements" rows="5" cols="80"><%=lcl_special_requirements%></textarea></td>
            </tr>
            <tr valign="top">
                <td>Misc Info:</td>
                <td colspan="4"><textarea name="misc_info" rows="5" cols="80"><%=lcl_misc_info%></textarea></td>
            </tr>
         <% if lcl_sc_list_type = "BID" then %>
            <tr>
                <td>Addendum Date:</td>
                <td colspan="4"><input type="text" name="bid_addendum_date" id="bid_addendum_date" value="<%=lcl_bid_addendum_date%>" size="15" maxlength="15" onchange="clearMsg('bid_addendum_date_pop');" />&nbsp;
                    <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" id="bid_addendum_date_pop" height="16" width="16" border="0" onclick="javascript:void doCalendar('bid_addendum_date');" /></span>
                    &nbsp;[mm/dd/yyyy]</td>
            </tr>
            <tr>
                <td>Pre-bid Meeting:</td>
                <td colspan="4"><input type="text" name="bid_pre_bid_meeting" value="<%=lcl_bid_pre_bid_meeting%>" size="50" maxlength="50"></td>
            </tr>
            <tr valign="top">
                <td>Contact Person:</td>
                <td colspan="4"><textarea name="bid_contact_person" rows="5" cols="80"><%=lcl_bid_contact_person%></textarea></td>
            </tr>
         <% end if %>
            <tr valign="top">
                <td>Download Available:</td>
                <td colspan="2">
                    <textarea id="download_available" name="download_available" rows="5" cols="80" onselect="storeCaret(this);" onclick="storeCaret(this);" onkeyup="storeCaret(this);" ondblclick="storeCaret(this);" maxlength="8000"><%=lcl_download_available%></textarea>
                </td>
                <td valign="top" style="padding-right:40px">
                <%
                  if lcl_orghasfeature_clickcounter_postings then
                     displayHelpIcon("clickcounter_postings")
                     response.write "<br />" & vbcrlf
                  end if
                %>
                				<input type="button" class="button" value="Add Link" onClick="doPicker('posting_maint.download_available','Y','Y','Y','Y');" />
                </td>
            </tr>
         <% if lcl_sc_list_type = "BID" then %>
            <tr>
                <td>Fee:</td>
                <td colspan="4"><input type="text" name="bid_fee" value="<%=lcl_bid_fee%>" size="50" maxlength="100" /></td>
            </tr>
            <tr valign="top">
                <td>Plan and Spec Available:</td>
                <td colspan="4"><textarea name="bid_plan_spec_available" rows="5" cols="80"><%=lcl_bid_plan_spec_available%></textarea></td>
            </tr>
            <tr>
                <td>Business Hours:</td>
                <td colspan="4"><input type="text" name="bid_business_hours" value="<%=lcl_bid_business_hours%>" size="50" maxlength="100" /></td>
            </tr>
            <tr>
                <td>Fax Number:</td>
                <td colspan="4"><input type="text" name="bid_fax_number" value="<%=lcl_bid_fax_number%>" size="15" maxlength="15" /></td>
            </tr>
            <tr valign="top">
                <td>Plan Holders:</td>
                <td colspan="4"><textarea name="bid_plan_holders" rows="5" cols="80"><%=lcl_bid_plan_holders%></textarea></td>
            </tr>
         <% end if %>
            <tr>
                <td align="center" colspan="4">
                    <fieldset>
                      <legend><%=lcl_label%> Posting Categories&nbsp;</legend>
                      <br />
                      <table border="0" cellspacing="1" cellpadding="2" width="100%" class="tableadmin">
                        <%
                         'Retrieve all of the top-level categories
                          sSQL = "SELECT distributionlistid, distributionlistname "
                          sSQL = sSQL & " FROM egov_class_distributionlist "
                          sSQL = sSQL & " WHERE orgid = " & session("orgid")
                          sSQL = sSQL & " AND (parentid = '' OR parentid IS NULL) "
                          sSQL = sSQL & " AND distributionlisttype = '" & lcl_sc_list_type & "'"
                          sSQL = sSQL & " ORDER BY UPPER(distributionlistname) "

                          set oTopLevelCats = Server.CreateObject("ADODB.Recordset")
                          oTopLevelCats.Open sSQL, Application("DSN") , 3, 1

                          lcl_bgcolor = "#ffffff"

                          if not oTopLevelCats.eof then
                             do while not oTopLevelCats.eof
                                lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")

                                response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf

                                lcl_checked = checkForCheckboxValue(oTopLevelCats("distributionlistid"),lcl_posting_id)

                                response.write "    <td align=""left"">" & vbcrlf
                                response.write "        <input type=""checkbox"" name=""category_" & oTopLevelCats("distributionlistid") & """ value=""" & oTopLevelCats("distributionlistid") & """" & lcl_checked & ">"

                                if lcl_sc_list_type = "BID" then
                                   response.write "        &nbsp;<b>" & oTopLevelCats("distributionlistname") & "</b>"
                                else
                                   response.write "        &nbsp;" & oTopLevelCats("distributionlistname")
                                end if

                                response.write "    </td>" & vbcrlf
                                response.write "</tr>" & vbcrlf

                               'Retrieve all of the sub-categories
                                sSQL = "SELECT distributionlistid, distributionlistname "
                                sSQL = sSQL & " FROM egov_class_distributionlist "
                                sSQL = sSQL & " WHERE orgid = " & session("orgid")
                                sSQL = sSQL & " AND parentid = " & oTopLevelCats("distributionlistid")
                                sSQL = sSQL & " AND distributionlisttype = '" & lcl_sc_list_type & "'"
                                sSQL = sSQL & " ORDER BY UPPER(distributionlistname) "

                                set oSubLevelCats = Server.CreateObject("ADODB.Recordset")
                                oSubLevelCats.Open sSQL, Application("DSN") , 3, 1

                                if not oSubLevelCats.eof then
                                   do while not oSubLevelCats.eof
                                      lcl_checked = checkForCheckboxValue(oSubLevelCats("distributionlistid"),lcl_posting_id)

                                      response.write "<tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
                                      response.write "    <td align=""left"">" & vbcrlf
                                      response.write "        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""checkbox"" name=""category_" & oSubLevelCats("distributionlistid") & """ value=""" & oSubLevelCats("distributionlistid") & """" & lcl_checked & ">"
                                      response.write "        &nbsp;" & oSubLevelCats("distributionlistname")
                                      response.write "    </td>" & vbcrlf
                                      response.write "</tr>" & vbcrlf
                                      oSubLevelCats.movenext
                                   loop
                                end if

                                oSubLevelCats.close
                                set oSubLevelCats = nothing

                                oTopLevelCats.movenext
                             loop
                          end if

                          oTopLevelCats.close
                          set oTopLevelCats = nothing
                        %>
                      </table>
                    </fieldset>
                </td>
            </tr>
          </table>
      </td>
  </tr>
  </form>
</table>
<!-- </div> -->
<% display_buttons lcl_sc_title, lcl_sc_status_id, lcl_sc_publicly_viewable, lcl_sc_list_type, lcl_sc_orderby, lcl_autosend_email %>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
'------------------------------------------------------------------------------
function display_buttons(p_sc_title, p_sc_status_id, p_sc_publicly_viewable, p_sc_list_type, p_sc_orderby, p_autosend_email)

  response.write "<div id=""functionlinks"">" & vbcrlf

 'Cancel
  'response.write "  <a href=""job_bid_list.asp?sc_title=" & p_sc_title & "&sc_status_id=" & lcl_sc_status_id & "&sc_publicly_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby & """>"
  'response.write "  <img src=""../images/cancel.gif"" align=""absmiddle"" border=""0"">&nbsp;Cancel</a>&nbsp;&nbsp;" & vbcrlf
  response.write "<input type=""button"" name=""cancelButton"" id=""cancelButton"" value=""Cancel"" class=""button"" onclick=""location.href='job_bid_list.asp?sc_title=" & p_sc_title & "&sc_status_id=" & lcl_sc_status_id & "&sc_publicly_viewable=" & lcl_sc_publicly_viewable & "&sc_list_type=" & lcl_sc_list_type & "&sc_orderby=" & lcl_sc_orderby & "';"" />" & vbcrlf

 'Save
  'response.write "  <a href=""javascript:document.posting_maint.submit();"">"
  'response.write "  <img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;" & sLinkText & "</a>&nbsp;&nbsp;" & vbcrlf
  'response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""" & sLinkText & """ class=""button"" onclick=""document.posting_maint.submit();"" />" & vbcrlf
  response.write "<input type=""button"" name=""saveButton"" id=""saveButton"" value=""" & sLinkText & """ class=""button"" onclick=""return validateFields();"" />" & vbcrlf

 'Save and Send Email
  'response.write "  <a href=""javascript:document.posting_maint.autosend_email.value='Y';document.posting_maint.submit();"">"
  'response.write "  <img src=""../images/go.gif"" align=""absmiddle"" border=""0"">&nbsp;" & sLinkText & " and Send E-mail</a>&nbsp;&nbsp;" & vbcrlf
  response.write "<input type=""button"" name=""saveSendEmailButton"" id=""saveSendEmailButton"" value=""" & sLinkText & " and Send Email"" class=""button"" onclick=""document.posting_maint.autosend_email.value='Y';document.posting_maint.submit();"" />" & vbcrlf

  response.write "</div>" & vbcrlf
end function

'------------------------------------------------------------------------------
function checkForCheckboxValue(p_distributionlistid, p_posting_id)
  if p_distributionlistid <> "" AND p_posting_id <> "" then
     sSQL = "SELECT ' checked' AS lcl_exists "
     sSQL = sSQL & " FROM egov_distributionlists_jobbids "
     sSQL = sSQL & " WHERE distributionlistid = " & p_distributionlistid
     sSQL = sSQL & " AND posting_id = " & p_posting_id

    	set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open sSQL, Application("DSN"), 3, 1

     if not rs.eof then
        lcl_checked = rs("lcl_exists")
     else
        lcl_checked = ""
     end if
  else
     lcl_checked = ""
  end if

  checkForCheckboxValue = lcl_checked

end function

'------------------------------------------------------------------------------
function checkForPermission(p_posting_id, p_sc_list_type)
  lcl_return = "Y"

 'Check the type of list, evaluate the screen mode, and then check for the permission
  if CLng(p_posting_id) = CLng(0) then
      if UCASE(p_sc_list_type) = "JOB" then
         if not UserHasPermission(session("userid"),"create_job_postings") then
  	         lcl_return = "N"
         end if
      elseif UCASE(p_sc_list_type) = "BID" then
         if not UserHasPermission(session("userid"),"create_bid_postings") then
  	         lcl_return = "N"
         end if
      end if
  else
      if UCASE(p_sc_list_type) = "JOB" then
         if not UserHasPermission(session("userid"),"edit_job_postings") then
  	         lcl_return = "N"
         end if
      elseif UCASE(p_sc_list_type) = "BID" then
         if not UserHasPermission(session("userid"),"edit_bid_postings") then
  	         lcl_return = "N"
         end if
      end if
  end if

  checkForPermission = lcl_return

end function

'------------------------------------------------------------------------------
sub displayPostingStatuses(p_orgid, p_sc_list_type, p_status_id)

 'Retreive all of the statuses for JOB Postings
  sSQL = "SELECT status_id, status_name "
  sSQL = sSQL & " FROM egov_statuses "
  sSQL = sSQL & " WHERE status_type = '" & p_sc_list_type & "' "
  sSQL = sSQL & " AND orgid = " & p_orgid
  sSQL = sSQL & " AND active_flag = 'Y' "
  sSQL = sSQL & " ORDER BY status_order "

  set oPostingStatus = Server.CreateObject("ADODB.Recordset")
  oPostingStatus.Open sSQL, Application("DSN") , 3, 1

  if not oPostingStatus.eof then
     do while not oPostingStatus.eof
        if clng(p_status_id) = clng(oPostingStatus("status_id")) then
           lcl_selected = " selected"
        else
           lcl_selected = ""
        end if

        response.write "  <option value=""" & oPostingStatus("status_id") & """" & lcl_selected & ">" & oPostingStatus("status_name") & "</option>" & vbcrlf

        oPostingStatus.movenext
     loop
  else
     response.write "  <option value=""0""></option>" & vbcrlf
  end if

  oPostingStatus.close
  set oPostingStatus = nothing

end sub
%>
