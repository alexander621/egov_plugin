<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME:  pool_attendance_maint.asp
' AUTHOR:    David Boyer
' CREATED:   04/28/08
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0  04/30/08  David Boyer - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
 if isFeatureOffline("memberships") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 sLevel     = "../"     'Override of value from common.asp
 lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=hide

'Determine if user is adding/editing
 if request("pid") = "" OR NOT isnumeric(request("pid")) OR request("pid") = 0 then
	  '-- ADD --------------------------------
	   lcl_poolinfo_id       = 0
    lcl_screen_mode_label = "Add"

    if checkForPermission(lcl_poolinfo_id) = "N" then
       response.redirect sLevel & "permissiondenied.asp"
    end if

    sTitle    = "Daily Pool Attendance: Add"
	   sLinkText = "Add"
 else
	  '-- EDIT -------------------------------
	   lcl_poolinfo_id       = request("pid")
    lcl_screen_mode_label = "Save"

    if checkForPermission(lcl_poolinfo_id) = "N" then
       response.redirect sLevel & "permissiondenied.asp"
    end if

    sTitle    = "Daily Pool Attendance: Maintain"
	   sLinkText = "Save"
 end if

'Retrieve data for this posting.
'	sSQL = "SELECT poolinfoid, orgid, pool_date, total_members, total_punchcards, total_guests, total_groups "
	sSQL = "SELECT poolinfoid, orgid, pool_date "
 sSQL = sSQL & " FROM egov_pool_info "
 sSQL = sSQL & " WHERE orgid = " & session("orgid")
 sSQL = sSQL & " AND poolinfoid = " & lcl_poolinfo_id

	set oValues = Server.CreateObject("ADODB.Recordset")
	oValues.Open sSQL, Application("DSN") , 3, 1

	If NOT oValues.EOF Then
    lcl_poolinfo_id = oValues("poolinfoid")
  		lcl_pool_date   = oValues("pool_date")
 else
    lcl_poolinfo_id = 0
  		lcl_pool_date   = date()
 end if

	oValues.close
	set oValues = nothing

'Determine which tab is active
 if request("activetab") <> "" then
   	iActiveTabId = clng(request("activetab"))
 else
   	iActiveTabId = clng(0)
 end if

 if request("tabid") <> "" then
   	iTabId = request("tabid")
 else
   	iTabId = "SAVE"
 end if
%>
<html>
<head>
	<title>E-Gov Administration Console - <%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />

	<!--
	<script type="text/javascript" src="../yui/build/yahoo-dom-event/yahoo-dom-event.js"></script>
	<script type="text/javascript" src="../yui/build/element/element-beta.js"></script>
	<script type="text/javascript" src="../yui/build/tabview/tabview.js"></script>
	-->
	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>

<script language="javascript">
		var tabView;
		var winHandle;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

 function SetUpPage(p_tabid) {
 	 tabView.set('activeIndex',p_tabid);
 }

 function doCalendar(sField) {
   var w = (screen.width - 350)/2;
   var h = (screen.height - 350)/2;
   eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=pool_maint", "_poolmaint", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
 }

 function doPicker(sFormField) {
   w = (screen.width - 350)/2;
   h = (screen.height - 350)/2;
   eval('window.open("../picker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
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

 function validateFields(p_tab_name) {
   if(p_tab_name=="WEATHER_ADD") {
      var rege = /^[0-9]{0,}$/;
      var lcl_new_air   = document.getElementById("p_new_weather_temp_air").value;
      var lcl_new_water = document.getElementById("p_new_weather_temp_water").value;

      if(lcl_new_air!="") {
         var result_air = rege.test(lcl_new_air);

         if(! result_air) {
            document.getElementById("p_new_weather_temp_air").focus();
            alert("Air Temperature must be numeric");
            return false;
         }else{
            result_air = true;
         }
      }

      if(result_air && lcl_new_water!="") {
         var result_water = rege.test(lcl_new_water);

         if(! result_water) {
            document.getElementById("p_new_weather_temp_water").focus();
            alert("Water Temperature must be numeric");
            return false;
         }
      }else{
         result_water = true;
      }

      return true;
   }else{
      lcl_active_tab = 0
      result_totals  = true;

      //Check to see if the user has modified the pool_date
      lcl_pooldate_orig = document.getElementById("pool_date_original").value;
      lcl_pooldate      = document.getElementById("pool_date").value;

      if(lcl_pooldate_orig != lcl_pooldate) {
         input_box=confirm("Are you sure that you want to change the Attendance Date from \"" + lcl_pooldate_orig + "\" to \"" + lcl_pooldate + "\"?");

         if (input_box==true) {
             // Date change has been verified
         }else{
             input_box=confirm("Do you want to reset the Attendance Date to its original value of \"" + lcl_pooldate_orig + "\" and continue saving your changes?");

             if (input_box==true) {
                 document.getElementById("pool_date").value=lcl_pooldate_orig;
             }else{
                 // Date Change has been denied
                 document.getElementById("pool_date").focus();
                 return false;
             }
         }
      }

      if((result_totals)&&("0"!="<%=lcl_poolinfo_id%>")) {
         //Validate the WEATHER tab
         lcl_active_tab = 1;
         var lcl_air    = "";
         var lcl_water  = "";
         var lcl_total_count_weather = document.getElementById("p_total_weather").value;

         if(lcl_total_count_weather>0) {
            for (i = 1; i <= lcl_total_count_weather; i++) {
                 var rege = /^[0-9]{0,}$/;

                 //Check to see if the object exists
                 if(document.getElementById("p_temperature_air_"+i)) { 
                    lcl_air = document.getElementById("p_temperature_air_"+i).value;
                 }

                 //Check to see if the object exists
                 if(document.getElementById("p_temperature_water_"+i)) { 
                    lcl_water = document.getElementById("p_temperature_water_"+i).value;
                 }

                 if(lcl_air!="") {
                    var result_air = rege.test(lcl_air);

                    if(! result_air) {
                       //Set up the tab info
                       SetUpPage(lcl_active_tab);
                       document.getElementById('activetab').value=lcl_active_tab;
                       document.getElementById('tabid').value='SAVE';

                       //Display error message and set focus to field in error.
                       document.getElementById("p_temperature_air_"+i).focus();
                       alert("Air Temperature must be numeric");
                       return false;
                       break;
                    }else{
                       result_air = true;
                    }
                 }

                 if(result_air && lcl_water!="") {
                    var result_water = rege.test(lcl_water);

                    if(! result_water) {
                       //Set up the tab info
                       SetUpPage(lcl_active_tab);
                       document.getElementById('activetab').value=lcl_active_tab;
                       document.getElementById('tabid').value='SAVE';

                       //Display error message and set focus to field in error.
                       document.getElementById("p_temperature_water_"+i).focus();
                       alert("Water Temperature must be numeric");
                       return false;
                       break;
                    }else{
                       result_water = true;
                    }
                 }
            }
         }
      }
      return true;
   }
}

function deleteconfirm(p_tabid,p_rowid) {
  if(p_tabid=="WEATHER") {
   		var tbl           = document.getElementById("weather_edit");
     var lcl_id        = document.getElementById("p_weatherid_"+p_rowid).value;
     var lcl_msg_label = "weather";
     var lcl_msg_value = document.getElementById("p_weather_time_"+p_rowid).value;

  }else if(p_tabid=="INCIDENT") {
   		var tbl           = document.getElementById("incidents_edit");
     var lcl_id        = document.getElementById("p_incidentid_"+p_rowid).value;
     var lcl_msg_label = "incident";
     var lcl_msg_value = document.getElementById("p_incident_time_hour_"+p_rowid).value;
         lcl_msg_value = lcl_msg_value + ":" + document.getElementById("p_incident_time_min_"+p_rowid).value;
         lcl_msg_value = lcl_msg_value + " " + document.getElementById("p_incident_time_ampm_"+p_rowid).value;
  }else if(p_tabid=="NOTE") {
   		var tbl           = document.getElementById("notes_edit");
     var lcl_id        = document.getElementById("p_noteid_"+p_rowid).value;
     var lcl_msg_label = "note";
     var lcl_msg_value = document.getElementById("note_datetime_"+p_rowid).innerHTML;
  }

  input_box=confirm("Are you sure you want to delete the \"" + lcl_msg_value + "\" " + lcl_msg_label + " record?");
  if (input_box==true) {
      // DELETE HAS BEEN VERIFIED
						doAjax('pool_attendance_action.asp', 'cmd=D&tabid=' + p_tabid + '&' + lcl_msg_label + 'id='+lcl_id, '', 'get', '0');
						tbl.deleteRow(p_rowid);
      document.getElementById("status_message").innerHTML = "<b style=\"color:#FF0000\">*** Successfully Deleted... ***</b>";
  }else{
      // CANCEL DELETE PROCESS
  }
}

function clearMsg() {
  document.getElementById("status_message").innerHTML = "&nbsp;";
}
</script>
</head>
<body class="yui-skin-sam" onload="SetUpPage(<%=iActiveTabId%>);">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="centercontent">
<table border="0" cellspacing="0" cellpadding="0" width="600" class="start">
  <tr>
      <td colspan="2">
          <font size="+1"><b><%=sTitle%></b></font><br>
<!--          <a href="pool_attendance_list.asp?use_sessions=Y"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to List</a> -->
          <p>
      </td>
  </tr>
  <tr>
      <td>
          <%
            if lcl_poolinfo_id = 0 then
               display_buttons "ADD"
            else
               display_buttons "SAVE"
            end if
          %>
      </td>
      <td id="status_message" align="right">
      <%
        lcl_message = ""

        if request("success") = "SU" then
           lcl_message = "<b style=""color:#FF0000"">*** Successfully Updated... ***</b>"
        elseif request("success") = "SA" then
           lcl_message = "<b style=""color:#FF0000"">*** Successfully Created... ***</b>"
        elseif request("success") = "AE" then
           lcl_message = "<b style=""color:#FF0000"">*** A record with this date already exists... ***</b>"
        else
           lcl_message = "&nbsp;"
        end if

        if lcl_message <> "" then
           response.write lcl_message
        end if
      %>
      </td>
  </tr>
</table>
<form name="pool_maint" id="pool_maint" action="pool_attendance_action.asp" method="post">
  <input type="<%=lcl_hidden%>" name="poolinfoid" value="<%=lcl_poolinfo_id%>">
		<input type="<%=lcl_hidden%>" name="activetab" id="activetab" value="<%=iActiveTabId%>" size="2" maxlength="5">
  <input type="<%=lcl_hidden%>" name="tabid" id="tabid" value="<%=iTabId%>">
  <input type="<%=lcl_hidden%>" name="pool_date_original" id="pool_date_original" value="<%=lcl_pool_date%>" size="10" maxlength="10">
<div id="demo" class="yui-navset">
    <ul class="yui-nav">
		    <li onclick="document.getElementById('activetab').value=0;document.getElementById('tabid').value='SAVE';"><a href="#tab1"><em id="tab0">Details</em></a></li>
      <% if lcl_poolinfo_id > 0 then %>
						<li onclick="document.getElementById('activetab').value=1;document.getElementById('tabid').value='SAVE'"><a href="#tab2"><em>Weather</em></a></li>
						<li onclick="document.getElementById('activetab').value=2;document.getElementById('tabid').value='SAVE'"><a href="#tab3"><em>Incidents</em></a></li>
						<li onclick="document.getElementById('activetab').value=3;document.getElementById('tabid').value='SAVE'"><a href="#tab4"><em>Notes</em></a></li>
      <% end if %>
					</ul>            
					<div class="yui-content">
<!-- DETAILS -->
     				<div id="tab0">
         <% getTotalMembers lcl_poolinfo_id, lcl_total_members, lcl_total_punchcards, lcl_total_guests, lcl_total_groups %>

             <table border="0" cellspacing="0" cellpadding="5">
               <tr>
                   <td>Attendance Date:</td>
                   <td>
                       <input type="text" name="pool_date" id="pool_date" value="<%=lcl_pool_date%>" size="10" maxlength="10">&nbsp;
                       <span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('pool_date');" /></span>
                   </td>
               </tr>
               <tr>
                   <td>Total Members:</td>
                   <td style="color: #800000"><%=lcl_total_members%></td>
               </tr>
               <tr>
                   <td>Total Punchcards:</td>
                   <td style="color: #800000"><%=lcl_total_punchcards%></td>
               </tr>
               <tr>
                   <td>Total Guests/Daily Rate:</td>
                   <td style="color: #800000"><%=lcl_total_guests%></td>
               </tr>
               <tr>
                   <td>Total Group Rate:</td>
                   <td style="color: #800000"><%=lcl_total_groups%></td>
               </tr>
               <tr>
                   <td><b>Overall Attendance:</b></td>
                   <td style="color: #800000"><b><%=lcl_total_members+lcl_total_punchcards+lcl_total_guests+lcl_total_groups%></b></td>
               </tr>
             </table>
         </div>

<% if lcl_poolinfo_id > 0 then %>
 <!-- WEATHER -->
     				<div id="tab1">
             <fieldset>
               <legend>Add Weather Info&nbsp;</legend>
               <table border="0" cellspacing="0" cellpadding="5">
                 <tr>
                     <td>
                         <table border="0" cellspacing="0" cellpadding="2">
                           <tr>
                               <td>Time:</td>
                               <td>
                                   <select name="p_new_weather_time">
                                     <%
                                        lcl_weather_hour = hour(time())

                                        if lcl_weather_hour > 12 then
                                           lcl_weather_hour = lcl_weather_hour - 12
                                        end if

                                        if lcl_weather_hour < 10 then
                                           lcl_weather_hour = "0" & lcl_weather_hour
                                        end if

                                        showTimeOptions lcl_weather_hour & ":00 " & right(time(),2)
                                     %>
                                   </select>
                               </td>
                           </tr>
                           <tr>
                               <td>Air Temp:</td>
                               <td><input type="text" name="p_new_weather_temp_air" id="p_new_weather_temp_air" size="6" maxlength="3"></td>
                           </tr>
                           <tr>
                               <td>Water Temp:</td>
                               <td><input type="text" name="p_new_weather_temp_water" id="p_new_weather_temp_water" size="6" maxlength="3"></td>
                           </tr>
                         </table>
                     </td>
                     <td valign="top">
                         <table border="0" cellspacing="0" cellpadding="2">
                           <tr valign="top">
                               <td>Description:</td>
                               <td><textarea name="p_new_weather_description" rows="5" cols="60"></textarea></td>
                           </tr>
                         </table>
                     </td>
                 </tr>
                 <tr>
                     <td colspan="2"><% display_buttons "WEATHER_ADD" %></td>
                 </tr>
               </table>
             </fieldset>
             <p>
             <table border="0" cellspacing="0" cellpadding="5" width="100%" id="weather_edit" class="tableadmin">
               <tr valign="bottom">
                   <th align="left">Time</th>
                   <th>Air<br>Temperature</th>
                   <th>Water<br>Temperature</th>
                   <th width="50%" align="left">Description</th>
                   <th>Delete</th>
               </tr>
               <%
                'Retrieve all of the weather records
                 sSQLw = "SELECT weatherid, poolinfoid, weather_time, temperature_air, temperature_water, description "
                 sSQLw = sSQLw & " FROM egov_pool_weather_log "
                 sSQLw = sSQLw & " WHERE poolinfoid = " & lcl_poolinfo_id
                 sSQLw = sSQLw & " AND orgid = " & session("orgid")
                 sSQLw = sSQLw & " ORDER BY RIGHT(weather_time,2), LEFT(weather_time,2) "

                 set rsw = Server.CreateObject("ADODB.Recordset")
                	rsw.Open sSQLw, Application("DSN") , 3, 1

                 lcl_total_count     = 0
                 lcl_bgcolor_weather = "#eeeeee"

                 if not rsw.eof then
                    while not rsw.eof
                       lcl_total_count = lcl_total_count + 1
                       lcl_bgcolor_weather = changeBGColor(lcl_bgcolor_weather,"","")

                       response.write "<tr id=""weather_row"" align=""center"" valign=""top"" bgcolor=""" & lcl_bgcolor_weather & """>" & vbcrlf
                       response.write "    <td align=""left"">"
                       response.write "        <input type=""" & lcl_hidden & """ name=""p_weatherid_" & lcl_total_count & """ id=""p_weatherid_" & lcl_total_count & """ value=""" & rsw("weatherid") & """ size=""5"" maxlength=""5"">" & vbcrlf
                       response.write "        <select name=""p_weather_time_" & lcl_total_count & """ id=""p_weather_time_" & lcl_total_count & """>" & vbcrlf
                                                 showTimeOptions rsw("weather_time")
                       response.write "        </select>" & vbcrlf
                       response.write "    </td>" & vbcrlf
                       response.write "    <td><input type=""text"" name=""p_temperature_air_"         & lcl_total_count & """ id=""p_temperature_air_"     & lcl_total_count & """ value="""   & rsw("temperature_air")   & """ size=""3"" maxlength=""3""></td>" & vbcrlf
                       response.write "    <td><input type=""text"" name=""p_temperature_water_"       & lcl_total_count & """ id=""p_temperature_water_"   & lcl_total_count & """ value=""" & rsw("temperature_water") & """ size=""3"" maxlength=""3""></td>" & vbcrlf
                       response.write "    <td align=""left""><textarea name=""p_weather_description_" & lcl_total_count & """ id=""p_weather_description_" & lcl_total_count & """ rows=""4"" cols=""60"">" & rsw("description")& "</textarea></td>" & vbcrlf
'                       response.write "    <td><input type=""checkbox"" name=""p_weather_delete_"      & lcl_total_count & """ value=""Y""></td>" & vbcrlf
                       response.write "    <td><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""clearMsg();deleteconfirm('WEATHER'," & lcl_total_count & ")""></td>" & vbcrlf
                       response.write "</tr>" & vbcrlf

                       rsw.movenext
                    wend
                 else
                       response.write "<tr><td colspan=""4"">No Records Exist</td></tr>" & vbcrlf
                 end if
               %>
               <tr><td colspan="5"><input type="<%=lcl_hidden%>" name="p_total_weather" id="p_total_weather" value="<%=lcl_total_count%>" size="5" maxlength="5"></td></tr>
             </table>
         </div>

<!-- INCIDENTS -->
     				<div id="tab2">
             <fieldset>
               <legend>Add Incident&nbsp;</legend>
               <table border="0" cellspacing="0" cellpadding="5">
                 <tr>
                     <td>
                         <table border="0" cellspacing="0" cellpadding="2">
                           <tr>
                               <td nowrap="nowrap">Incident Time:</td>
                               <td>
                                   <select name="p_new_incident_time_hour">
                                     <% buildHourMinOptions "HOURS",hour(time()) %>
                                   </select>
                                   :
                                   <select name="p_new_incident_time_min">
                                     <% buildHourMinOptions "MINUTES",minute(time()) %>
                                   </select>
                                   <select name="p_new_incident_time_ampm">
                                     <%
                                       if right(time(),2) = "PM" then
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
                           </tr>
                           <tr>
                               <td nowrap="nowrap">Name of Injured:</td>
                               <td><input type="text" name="p_new_incident_nameofinjured" size="30" maxlength="500"></td>
                           </tr>
                           <tr>
                               <td nowrap="nowrap">Type of Injury:</td>
                               <td><input type="text" name="p_new_incident_typeofinjury" size="30" maxlength="500"></td>
                           </tr>
                         </table>
                     </td>
                     <td valign="top">
                         <table border="0" cellspacing="0" cellpadding="2">
                           <tr>
                               <td nowrap="nowrap">Witness:</td>
                               <td><input type="text" name="p_new_incident_witness" size="30" maxlength="500"></td>
                           </tr>
                           <tr>
                               <td nowrap="nowrap">Report Completed By:</td>
                               <td><input type="text" name="p_new_incident_completedby" size="30" maxlength="500"></td>
                           </tr>
                         </table>
                     </td>
                 </tr>
                 <tr valign="top">
                     <td colspan="2">
                         <table border="0" cellspacing="0" cellpadding="2">
                           <tr valign="top">
                               <td nowrap="nowrap">Staff Response:</td>
                               <td><textarea name="p_new_incident_staffresponse" rows="3" cols="90"></textarea></td>
                           </tr>
                         </table>
                     </td>
                 </tr>
                 <tr>
                     <td colspan="2"><% display_buttons "INCIDENT_ADD" %></td>
                 </tr>
               </table>
             </fieldset>
             <p>
             <table border="0" cellspacing="0" cellpadding="2" width="100%" id="incidents_edit" class="tableadmin">
               <tr>
                   <th>Incident Time</th>
                   <th>Incident Details</th>
                   <th>Delete</th>
               </tr>
               <%
                'Retrieve all of the incident records
                 sSQLi = "SELECT incidentid, poolinfoid, incident_time, incident_time_ampm, name_of_injured, injury_type, witness, "
                 sSQLi = sSQLi & " staff_response, report_completed_by, report_completed_by_datetime "
                 sSQLi = sSQLi & " FROM egov_pool_incidents_log "
                 sSQLi = sSQLi & " WHERE poolinfoid = " & lcl_poolinfo_id
                 sSQLi = sSQLi & " AND orgid = " & session("orgid")
                 sSQLi = sSQLi & " ORDER BY incident_time_ampm, CAST(REPLACE(LEFT(incident_time,2),':','') AS INT), CAST(REPLACE(RIGHT(incident_time,2),':','') AS INT) "

                 set rsi = Server.CreateObject("ADODB.Recordset")
                	rsi.Open sSQLi, Application("DSN") , 3, 1

                 lcl_total_count      = 0
                 lcl_bgcolor_incident = "#ffffff"

                 if not rsi.eof then
                    while not rsi.eof
                       lcl_total_count      = lcl_total_count + 1
                       lcl_bgcolor_incident = changeBGColor(lcl_bgcolor_incident,"","")

                       response.write "<tr id=""incident_row"" align=""center"" valign=""top"" bgcolor=""" & lcl_bgcolor_incident & """>" & vbcrlf
                       response.write "    <td>" & vbcrlf
                       response.write "        <input type=""" & lcl_hidden & """ name=""p_incidentid_" & lcl_total_count & """ id=""p_incidentid_" & lcl_total_count & """ value=""" & rsi("incidentid") & """ size=""5"" maxlength=""5"">" & vbcrlf
                       response.write "        <select name=""p_incident_time_hour_" & lcl_total_count & """ id=""p_incident_time_hour_" & lcl_total_count & """>" & vbcrlf
                                                 buildHourMinOptions "HOURS",LEFT(rsi("incident_time"),2)
                       response.write "        </select>" & vbcrlf
                       response.write "        : " & vbcrlf
                       response.write "        <select name=""p_incident_time_min_" & lcl_total_count & """ id=""p_incident_time_min_" & lcl_total_count & """>" & vbcrlf
                                                 buildHourMinOptions "MINUTES",RIGHT(rsi("incident_time"),2)
                       response.write "        </select>" & vbcrlf
                       response.write "        <select name=""p_incident_time_ampm_" & lcl_total_count & """ id=""p_incident_time_ampm_" & lcl_total_count & """>" & vbcrlf

                       if rsi("incident_time_ampm") = "PM" then
                          lcl_selected_am = ""
                          lcl_selected_pm = " selected"
                       else
                          lcl_selected_am = " selected"
                          lcl_selected_pm = ""
                       end if

                       response.write "          <option value=""AM""" & lcl_selected_am & ">AM</option>" & vbcrlf
                       response.write "          <option value=""PM""" & lcl_selected_pm & ">PM</option>" & vbcrlf
                       response.write "        </select>" & vbcrlf
                       response.write "    </td>" & vbcrlf
                       response.write "    <td>" & vbcrlf
                       response.write "        <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
                       response.write "          <tr align=""left"" valign=""top"">" & vbcrlf
                       response.write "              <td><b>Name of Injured:</b></td>" & vbcrlf
                       response.write "              <td><input type=""text"" name=""p_name_of_injured_" & lcl_total_count & """ value=""" & rsi("name_of_injured") & """ size=""25"" maxlength=""500""></td>" & vbcrlf
                       response.write "              <td><b>Witness:</b></td>" & vbcrlf
                       response.write "              <td><input type=""text"" name=""p_witness_" & lcl_total_count & """ value=""" & rsi("witness") & """ size=""20"" maxlength=""500""></td>" & vbcrlf
                       response.write "          </tr>" & vbcrlf
                       response.write "          <tr align=""left"" valign=""top"">" & vbcrlf
                       response.write "              <td><b>Type of Injury:</b></td>" & vbcrlf
                       response.write "              <td><input type=""text"" name=""p_injury_type_" & lcl_total_count & """ value=""" & rsi("injury_type") & """ size=""25"" maxlength=""500""></td>" & vbcrlf
                       response.write "              <td><b>Report Completed By:</b></td>" & vbcrlf
                       response.write "              <td><input type=""text"" name=""p_report_completed_by_" & lcl_total_count & """ value="""   & rsi("report_completed_by") & """ size=""20"" maxlength=""500""></td>" & vbcrlf
                       response.write "          </tr>" & vbcrlf
                       response.write "          <tr align=""left"" valign=""top"">" & vbcrlf
                       response.write "              <td><b>Staff Response:</b></td>" & vbcrlf
                       response.write "              <td>&nbsp;</td>" & vbcrlf
                       response.write "              <td><b>Report Completed Date:</b></td>" & vbcrlf
                       response.write "              <td style=""color: #800000"">" & rsi("report_completed_by_datetime") & "</td>" & vbcrlf
                       response.write "          </tr>" & vbcrlf
                       response.write "          <tr align=""left"" valign=""top"">" & vbcrlf
                       response.write "              <td colspan=""4""><textarea name=""p_staff_response_" & lcl_total_count & """ rows=""3"" cols=""90"">" & rsi("staff_response")& "</textarea></td>" & vbcrlf
                       response.write "          </tr>" & vbcrlf
                       response.write "        </table>" & vbcrlf
                       response.write "    </td>" & vbcrlf
'                       response.write "    <td><input type=""checkbox"" name=""p_incident_delete_" & lcl_total_count & """ value=""Y""></td>" & vbcrlf
                       response.write "    <td><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""clearMsg();deleteconfirm('INCIDENT'," & lcl_total_count & ")""></td>" & vbcrlf
                       response.write "</tr>" & vbcrlf

                       rsi.movenext
                    wend
                 else
                       response.write "<tr><td colspan=""4"">No Records Exist</td></tr>" & vbcrlf
                 end if
               %>
               <tr><td colspan="5"><input type="<%=lcl_hidden%>" name="p_total_incidents" value="<%=lcl_total_count%>" size="5" maxlength="5"></td></tr>
             </table>
         </div>

<!-- NOTES -->
     				<div id="tab3">
             <fieldset>
               <legend>Add Note&nbsp;</legend>
               <table border="0" cellspacing="0" cellpadding="5">
                 <tr valign="top">
                     <td colspan="2">
                         <table border="0" cellspacing="0" cellpadding="2">
                           <tr valign="top">
                               <td nowrap="nowrap">Submitted By:</td>
                               <td><input type="text" name="p_new_submittedby" size="30" maxlength="500"></td>
                           </tr>
                           <tr valign="top">
                               <td nowrap="nowrap">Note:</td>
                               <td><textarea name="p_new_note" rows="3" cols="90"></textarea></td>
                           </tr>
                         </table>
                     </td>
                 </tr>
                 <tr>
                     <td colspan="2"><% display_buttons "NOTES_ADD" %></td>
                 </tr>
               </table>
             </fieldset>
             <p>
             <table border="0" cellspacing="0" cellpadding="5" width="100%" id="notes_edit" class="tableadmin">
               <tr valign="bottom">
                   <th>Submitted Date</th>
                   <th align="left">Submitted By</th>
                   <th width="45%" align="left">Note</th>
                   <th>Delete</th>
               </tr>
               <%
                'Retrieve all of the note records
                 sSQLn = "SELECT noteid, poolinfoid, note_submittedby, note_datetime, description "
                 sSQLn = sSQLn & " FROM egov_pool_info_notes "
                 sSQLn = sSQLn & " WHERE poolinfoid = " & lcl_poolinfo_id
                 sSQLn = sSQLn & " AND orgid = " & session("orgid")
                 sSQLn = sSQLn & " ORDER BY note_datetime DESC "

                 set rsn = Server.CreateObject("ADODB.Recordset")
                	rsn.Open sSQLn, Application("DSN") , 3, 1

                 lcl_total_count   = 0
                 lcl_bgcolor_notes = "#ffffff"

                 if not rsn.eof then
                    while not rsn.eof
                       lcl_total_count   = lcl_total_count + 1
                       lcl_bgcolor_notes = changeBGColor(lcl_bgcolor_notes,"","")

                       response.write "<tr id=""note_row"" align=""center"" valign=""top"" bgcolor=""" & lcl_bgcolor_notes & """>" & vbcrlf
                       response.write "    <td id=""note_datetime_" & lcl_total_count & """>" & rsn("note_datetime") & "</td>" & vbcrlf
                       response.write "    <td align=""left"">"
                       response.write "        <input type=""" & lcl_hidden & """ name=""p_noteid_" & lcl_total_count & """ id=""p_noteid_" & lcl_total_count & """ value=""" & rsn("noteid") & """ size=""5"" maxlength=""5"">" & vbcrlf
                       response.write "        <input type=""text"" name=""p_note_submittedby_" & lcl_total_count & """ id=""p_submittedby_" & lcl_total_count & """ value=""" & rsn("note_submittedby") & """ size=""30"" maxlength=""500"">" & vbcrlf
                       response.write "    </td>" & vbcrlf
                       response.write "    <td align=""left""><textarea name=""p_notes_description_" & lcl_total_count & """ id=""p_notes_description_" & lcl_total_count & """ rows=""4"" cols=""50"">" & rsn("description")& "</textarea></td>" & vbcrlf
'                       response.write "    <td><input type=""checkbox"" name=""p_notes_delete_" & lcl_total_count & """ value=""Y""></td>" & vbcrlf
                       response.write "    <td><img src=""../images/small_delete.gif"" border=""0"" alt=""Click to delete"" style=""cursor: hand"" onclick=""clearMsg();deleteconfirm('NOTE'," & lcl_total_count & ")""></td>" & vbcrlf
                       response.write "</tr>" & vbcrlf

                       rsn.movenext
                    wend
                 else
                       response.write "<tr><td colspan=""4"">No Records Exist</td></tr>" & vbcrlf
                 end if
               %>
               <tr><td colspan="5"><input type="<%=lcl_hidden%>" name="p_total_notes" id="p_total_notes" value="<%=lcl_total_count%>" size="5" maxlength="5"></td></tr>
             </table>
         </div>
<%
  '-----------------------------------------------------------
   end if
  '-----------------------------------------------------------
%>
      </td>
  </tr>
</table>
</form>
</div>  <%' end if TABs DIV %>

</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>
<%
function display_buttons(p_section)
  response.write "<div id=""functionlinks"">" & vbcrlf

  if UCASE(p_section) = "WEATHER_ADD" then
'     response.write "  <input type=""button"" name=""sAction"" value=""Add Weather Record"" onclick=""if(validateFields('WEATHER_ADD')){document.weather_maint_add.submit();}"">" & vbcrlf
     response.write "  <input type=""button"" name=""sAction"" value=""Add Weather Record"" onclick=""clearMsg();if(validateFields('" & UCASE(p_section) & "')){document.getElementById('tabid').value='" & UCASE(p_section) & "';document.pool_maint.submit();}"">" & vbcrlf
  elseif UCASE(p_section) = "INCIDENT_ADD" then
     response.write "  <input type=""button"" name=""sAction"" value=""Add Incident"" onclick=""clearMsg();document.getElementById('tabid').value='" & UCASE(p_section) & "';document.pool_maint.submit();"">" & vbcrlf
  elseif UCASE(p_section) = "NOTES_ADD" then
     response.write "  <input type=""button"" name=""sAction"" value=""Add Note"" onclick=""clearMsg();if(validateFields('" & UCASE(p_section) & "')){document.getElementById('tabid').value='" & UCASE(p_section) & "';document.pool_maint.submit();}"">" & vbcrlf
  elseif UCASE(p_section) = "ADD" then
     response.write "  <input type=""button"" id=""save_button"" value=""Add"" onclick=""clearMsg();if(validateFields('" & UCASE(p_section) & "')){document.pool_maint.submit();}"">" & vbcrlf
'  elseif UCASE(p_section) = "WEATHER_EDIT" then
'     response.write "  <input type=""button"" name=""sAction"" value=""Save"" onclick=""if(validateFields('WEATHER_EDIT')){document.weather_maint_edit.submit();}"">" & vbcrlf
'     response.write "  <input type=""button"" name=""sAction"" value=""Save"" onclick=""if(validateFields('WEATHER_EDIT')){document.pool_maint.submit();}"">" & vbcrlf
'  elseif UCASE(p_section) = "INCIDENT_EDIT" then
'     response.write "  <input type=""submit"" name=""sAction"" value=""Save"">" & vbcrlf
'  elseif UCASE(p_section) = "NOTES_EDIT" then
'     response.write "  <input type=""submit"" name=""sAction"" value=""Save"">" & vbcrlf
  else
    'Save
     response.write "  <input type=""button"" id=""save_button"" value=""Save Changes"" onclick=""clearMsg();if(validateFields('" & UCASE(p_section) & "')){document.pool_maint.submit();}"">" & vbcrlf
'     response.write "  <input type=""button"" value=""Save"" onclick=""if(validateFields('ALL')){document.pool_maint.submit();}"">" & vbcrlf

    'Cancel
'     response.write "  <input type=""button"" value=""Cancel"" onclick=""location.href='pool_attendance_list.asp?use_sessions=Y';"">" & vbcrlf
     response.write "  <input type=""button"" id=""return_button"" value=""Return to List"" onclick=""location.href='pool_attendance_list.asp?use_sessions=Y';"">" & vbcrlf
  end if

  response.write "</div>" & vbcrlf
end function

'------------------------------------------------------------------
function checkForPermission(p_poolinfoid)
  lcl_return = "Y"

 'Check the type of list, evaluate the screen mode, and then check for the permission
  if CLng(p_poolinfoid) = CLng(0) then
     if not UserHasPermission(session("userid"),"pool_attendance_add") then
        lcl_return = "N"
     end if
  else
     if not UserHasPermission(session("userid"),"pool_attendance_edit") then
        lcl_return = "N"
     end if
  end if

  checkForPermission = lcl_return

end function

'---------------------------------------------------------------
sub showTimeOptions(p_value)
   for i = 12 to 35
       if i - 12 < 1 then
          lcl_time = "12:00 AM"
       elseif i - 12 < 13 then
          if i - 12 < 10 then
             lcl_time = "0" & i - 12 & ":00 AM"
          else
             if i - 12 = 12 then
                lcl_time = i - 12 & ":00 PM"
             else
                lcl_time = i - 12 & ":00 AM"
             end if
          end if
       else
          if i - 24 < 10 then
             lcl_time = "0" & i - 24 & ":00 PM"
          else
             if i - 24 < 12 then
                lcl_time = i - 24 & ":00 PM"
             end if
          end if
       end if

       if p_value = lcl_time then
          lcl_selected = " selected"
       else
          lcl_selected = ""
       end if

       response.write "<option value=""" & lcl_time & """" & lcl_selected & ">" & lcl_time & "</option>" & vbcrlf
   next
end sub

'--------------------------------------------------------------------
sub buildHourMinOptions(p_dropdown_type,p_value)

 'Format the value passed in
 '1. Modify the hour to 1 - 12 (on NEW records the hour passed in is in 24hr time).
 '2. If the value is a single digit (less than 10) than concatenate a (0) to the front.
  if p_value > 12 then
     lcl_value = p_value - 12
  else
     lcl_value = p_value
  end if

  if lcl_value < 10 AND len(lcl_value) = 1 then
     lcl_value = "0" & lcl_value
  end if

  if UCASE(p_dropdown_type) = "MINUTES" then
     lcl_start = 0
     lcl_end   = 59
  else
     lcl_start = 1
     lcl_end   = 12
  end if

  for i = lcl_start to lcl_end
      if i < 10 then
         lcl_time_value = "0" & i
      else
         lcl_time_value = i
      end if

      if CStr(lcl_value) = CStr(lcl_time_value) then
         lcl_selected = " selected"
      else
         lcl_selected = ""
      end if

      response.write "<option value=""" & lcl_time_value & """" & lcl_selected & ">" & lcl_time_value & "</option>" & vbcrlf
  next
end sub

'--------------------------------------------------------------------
 sub getTotalMembers(ByVal p_poolinfo_id, ByRef lcl_total_members, ByRef lcl_total_punchcards, ByRef lcl_total_guests, ByRef lcl_total_groups )
  sSQL = "SELECT total_members, total_punchcards, total_guests, total_groups_peoplecount "
  sSQL = sSQL & " FROM egov_pool_info_vw "
  sSQL = sSQL & " WHERE poolinfoid = " & p_poolinfo_id

  set oTotals = Server.CreateObject("ADODB.Recordset")
  oTotals.Open sSQL, Application("DSN") , 3, 1

  if not oTotals.eof then
     lcl_total_members    = oTotals("total_members")
     lcl_total_punchcards = oTotals("total_punchcards")
     lcl_total_guests     = oTotals("total_guests")
     lcl_total_groups     = oTotals("total_groups_peoplecount")

     if lcl_total_members = "" OR isnull(lcl_total_members) then
        lcl_total_members = 0
     end if

     if lcl_total_punchcards = "" OR isnull(lcl_total_punchcards) then
        lcl_total_punchcards = 0
     end if

     if lcl_total_guests = "" OR isnull(lcl_total_guests) then
        lcl_total_guests = 0
     end if

     if lcl_total_groups = "" OR isnull(lcl_total_groups) then
        lcl_total_groups = 0
     end if

  else
     lcl_total_members    = 0
     lcl_total_punchcards = 0
     lcl_total_guests     = 0
     lcl_total_groups     = 0
  end if

  set oTotals = nothing

'  getTotalMembers = getTotalMembers

end sub

sub dtb_debug(p_value)
  sSQLi2 = "INSERT INTO my_table_dtb (notes) VALUES ('" & REPLACE(p_value,"'","''") & "')"
  set rsi2 = Server.CreateObject("ADODB.Recordset")
  rsi2.Open sSQLi2, Application("DSN") , 3, 1

end sub
%>
