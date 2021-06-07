<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: pool_attendance_list.asp
' AUTHOR:   David Boyer
' CREATED:  04/30/08
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  to allow users to maintain daily statistics about the pool
'
' MODIFICATION HISTORY
' 1.0  04/30/08  David Boyer - Created Code.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel     = "../" ' Override of value from common.asp
lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide

if NOT UserHasPermission( session("userid"), "pool_attendance_reporting" ) then
   response.redirect sLevel & "permissiondenied.asp"
end if

'Retrieve the search criteria fields
 if UCASE(request("use_sessions")) = "Y" then
    lcl_sc_from_date = session("sc_from_date")
    lcl_sc_to_date   = session("sc_to_date")
    lcl_sc_order_by  = session("sc_order_by")
 else
    lcl_sc_from_date = request("sc_from_date")
    lcl_sc_to_date   = request("sc_to_date")
    lcl_sc_order_by  = request("sc_order_by")
 end if

'Setup the session variables
 session("sc_from_date") = lcl_sc_from_date
 session("sc_to_date")   = lcl_sc_to_date
 session("sc_order_by")  = lcl_sc_order_by

 set oClassDivOrg = New classOrganization
%>
<html>
<head>
  <title>E-GovLink {Pool Attendance Reports}</title>
  
	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="reportprint.css" media="print" />

 <script src="../scripts/selectAll.js"></script>
 <script language="javascript" src="../scripts/modules.js"></script>

<script language="javascript">
		window.onload = function()
		{
		  //factory.printing.header = "Printed on &d"
    factory.printing.header       = "&bPrinted on &d"
		  factory.printing.footer       = "&bPrinted on &d - Page:&p/&P";
		  factory.printing.portrait     = true;
		  factory.printing.leftMargin   = 0.5;
		  factory.printing.topMargin    = 0.5;
		  factory.printing.rightMargin  = 0.5;
		  factory.printing.bottomMargin = 0.5;
		 
		  // enable control buttons
		  var templateSupported = factory.printing.IsTemplateSupported();
		  var controls = idControls.all.tags("input");
		  for ( i = 0; i < controls.length; i++ ) 
		  {
			controls[i].disabled = false;
			if ( templateSupported && controls[i].className == "ie55" )
			  controls[i].style.display = "inline";
		  }
		}

function doCalendar(ToFrom) {
  w = (screen.width - 350)/2;
  h = (screen.height - 350)/2;
  eval('window.open("calendarpicker.asp?p=1&ToFrom=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
}

function deleteconfirm(p_poolinfo_id) {
  lcl_pool_date = document.getElementById("pool_date_"+p_poolinfo_id).innerHTML;
  input_box=confirm("Are you sure you want to delete the \"" + lcl_pool_date + "\" attendance record?");
  if (input_box==true) {
      // DELETE HAS BEEN VERIFIED
      location.href='pool_attendance_list.asp?cmd=delete_poolinfoid&pid='+ p_poolinfo_id;
  }else{
      // CANCEL DELETE PROCESS
  }
}

checked=false;
function checkedAll () {
	 var x = document.getElementById('pool_list');
	 if (checked == false) {
      checked = true
  }else{
      checked = false
  }

 	for (var i=0; i < x.elements.length; i++) {
     	 x.elements[i].checked = checked;
 	}
}
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0"> <!-- onLoad="document.searchform.username.focus()"> -->

<div id="idControls" class="noprint">
	<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
	<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
</div>

<object id="factory" viewastext  style="display:none"
  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
   codebase="../includes/smsx.cab#Version=6,3,434,12">
</object>

<div id="content">
    <div id="centercontent">

<!--
<table border="0" cellpadding="0" cellspacing="0" class="start" width="100%">
  <form action="pool_reports.asp" method="post" name="pool_reports" id="pool_reports">
  <tr id="pool_report_links">
      <td><a href="pool_attendance_list.asp?use_sessions=Y"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to List</a><p></td>
  </tr>
  <tr>
      <td valign="top">
-->
<!--
<div id="pool_report_links">
  <a href="pool_attendance_list.asp?use_sessions=Y"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return to List</a><p>
</div>
-->
<input type="button" id="return_button" value="Return to List" onclick="location.href='pool_attendance_list.asp?use_sessions=Y';">
<p>

<%
 'Determine the report type
  if request("report_type") <> "" then
     lcl_report_type = request("report_type")
  else
     lcl_report_type = "DAILY"
  end if

 'Retrieve the records to be searched on
  if request("total_records") <> "" then
     lcl_total_count = request("total_records")
  else
     lcl_total_count = 0
  end if

  if lcl_total_count > 0 then
     lcl_poolinfo_ids = ""
     for x = 1 to lcl_total_count
         if request("checkbox_" & x) <> "" then
            if lcl_poolinfo_ids <> "" then
               lcl_poolinfo_ids = lcl_poolinfo_ids & "," & request("checkbox_" & x)
            else
               lcl_poolinfo_ids = request("checkbox_" & x)
            end if
         else
            lcl_poolinfo_ids = lcl_poolinfo_ids
         end if
     next
  else
     lcl_poolinfo_ids = 0
  end if

  'Build the reports
  if UCASE(lcl_report_type) = "DAILY" then
     for x = 1 to lcl_total_count
        if request("checkbox_" & x) <> "" then
          'All pages but the first will need to start on their own page.
           if x > 1 then
              response.write "<div class=""page_start"">" & vbcrlf
           end if

           displayReportHeader(request("checkbox_"&x))
           displayAttendanceTotals(request("checkbox_"&x))
           displayAttendanceTotalsPerHour(request("checkbox_"&x))
           displayWeatherAverages(request("checkbox_"&x))
           displayIncidents(request("checkbox_"&x))
           displayNotes(request("checkbox_"&x))
           displayFooter()

           if x > 1 then
              response.write "</div>" & vbcrlf
           end if
        end if
     next

  elseif UCASE(lcl_report_type) = "SUMMARY" then
     if lcl_poolinfo_ids <> "" then
        displayReportHeader(lcl_poolinfo_ids)
        displayAttendanceTotals(lcl_poolinfo_ids)
        displayAttendanceTotalsPerHour(lcl_poolinfo_ids)
        displayWeatherAverages(lcl_poolinfo_ids)
        displayFooter()
     end if

  elseif UCASE(lcl_report_type) = "INCIDENT" then
     if lcl_poolinfo_ids <> "" then
        displayReportHeader(lcl_poolinfo_ids)
        displayIncidents(lcl_poolinfo_ids)
        displayFooter()
     end if
  end if
%>
<!--
      </td>
  </tr>
  </form>
</table>
-->
    </div>
</div>

</body>
</html>
<%
'----------------------------------------------------------------------
 sub displayReportHeader(p_poolinfo_ids)
    'Retrieve the date(s)
     if instr(p_poolinfo_ids,",") > 0 then
        sSQL = "SELECT min(pool_date) AS min_date, max(pool_date) AS max_date " 
        sSQL = sSQL & " FROM egov_pool_info_vw "
        sSQL = sSQL & " WHERE poolinfoid IN (" & p_poolinfo_ids & ") "

        set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sSQL, Application("DSN"), 3, 1

        if not rs.eof then
           lcl_date_header = rs("min_date") & " - " & rs("max_date")
        else
           lcl_date_header = ""
        end if
     else
        sSQL = "SELECT pool_date " 
        sSQL = sSQL & " FROM egov_pool_info_vw "
        sSQL = sSQL & " WHERE poolinfoid = " & p_poolinfo_ids

        set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sSQL, Application("DSN"), 3, 1

        if not rs.eof then
           lcl_date_header = rs("pool_date")
        else
           lcl_date_header = ""
        end if
     end if

     response.write "<div>" & vbcrlf
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"">" & vbcrlf
   		response.write "  <tr>" & vbcrlf
     response.write "      <td align=""center"">" & vbcrlf
     response.write "          <font size=""+1""><strong>" & session("sOrgName") & "&nbsp;Community Pool Visit Report</strong></font>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr><td align=""center"">" & lcl_date_header & "</td></tr>" & vbcrlf
     response.write "</table>" & vbcrlf
     response.write "</div>" & vbcrlf

 end sub

'--------------------------------------------------------------------
 sub displayFooter()
   response.write "<p>" & vbcrlf
			response.write "<div class=""footerbox"">" & vbcrlf
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
			response.write "  <tr><td height=""5"" bgcolor=""#93bee1"" style=""border-bottom: solid 1px #000000;"">&nbsp;</td></tr>" & vbcrlf
			response.write "  <tr>" & vbcrlf
			response.write "      <td valign=""top"" align=""center"">" & vbcrlf
			response.write "    						<font style=""font-size:10px;font-weight:bold;"">Copyright &copy;2004-" & year(date()) & ".  "
   response.write "          All Rights Reserved. " & oClassDivOrg.GetOrgDisplayName( "admin footer brand link" )
 		response.write "          </font><p>" & vbcrlf
			response.write "						</td>" & vbcrlf
			response.write "		</tr>" & vbcrlf
			response.write "</table>" & vbcrlf
			response.write "</div>" & vbcrlf
 end sub

'-------------------------------------------------------------------
 sub displayAttendanceTotals(p_poolinfo_ids)
   response.write "<p>" & vbcrlf
   response.write "<div align=""left"">" & vbcrlf
	 	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""width: 300px"">" & vbcrlf
   response.write "  <caption align=""left""><b>Attendance</b></caption>" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <th width=""80%"">Type</th>" & vbcrlf
   response.write "      <th align=""center"">Total</th>" & vbcrlf
   response.write "  </tr>" & vbcrlf

  'Retrieve all of the totals
   sSQL = "SELECT SUM(total_members) as overall_members, "
   sSQL = sSQL & " SUM(total_punchcards) as overall_punchcards, "
   sSQL = sSQL & " SUM(total_guests) as overall_guests, "
   sSQL = sSQL & " SUM(total_groups_peoplecount) as overall_groups "
   sSQL = sSQL & " FROM egov_pool_info_vw "
   sSQL = SSQL & " WHERE poolinfoid in (" & p_poolinfo_ids & ")"

   set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open sSQL, Application("DSN"), 3, 1

   if not rs.eof then
      lcl_overall_members    = 0
      lcl_overall_punchcards = 0
      lcl_overall_guests     = 0
      lcl_overall_groups     = 0

      if rs("overall_members") <> "" then
         lcl_overall_members = rs("overall_members")
      end if

      if rs("overall_punchcards") <> "" then
         lcl_overall_punchcards = rs("overall_punchcards")
      end if

      if rs("overall_guests") <> "" then
         lcl_overall_guests = rs("overall_guests")
      end if

      if rs("overall_groups") <> "" then
         lcl_overall_groups = rs("overall_groups")
      end if

      response.write "  <tr>" & vbcrlf
      response.write "      <td>Members</td>" & vbcrlf
      response.write "      <td align=""center"">" & lcl_overall_members & "</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
      response.write "  <tr>" & vbcrlf
      response.write "      <td>Punchcards</td>" & vbcrlf
      response.write "      <td align=""center"">" & lcl_overall_punchcards & "</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
      response.write "  <tr>" & vbcrlf
      response.write "      <td>Guests/Daily Rate</td>" & vbcrlf
      response.write "      <td align=""center"">" & lcl_overall_guests & "</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
      response.write "  <tr>" & vbcrlf
      response.write "      <td>Groups</td>" & vbcrlf
      response.write "      <td align=""center"">" & lcl_overall_groups & "</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf

      set rs = nothing

   end if

   response.write "</table>" & vbcrlf
			response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""width: 300px"">" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <td align=""right"" width=""80%"">" & vbcrlf
   response.write "          <strong>Overall Total: </srong>" & vbcrlf
   response.write "      </td>" & vbcrlf
   response.write "      <td align=""center"">" & vbcrlf
   response.write "          <strong>" & lcl_overall_members + lcl_overall_punchcards + lcl_overall_guests + lcl_overall_groups & "</strong>" & vbcrlf
   response.write "      </td>" & vbcrlf
   response.write "  </tr>" & vbcrlf
   response.write "</table>" & vbcrlf
   response.write "</div>" & vbcrlf
 end sub

'--------------------------------------------------------------------------
sub displayAttendanceTotalsPerHour(p_poolinfo_ids)
   response.write "<p>" & vbcrlf
   response.write "<div align=""left"">" & vbcrlf
	 	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""width: 300px"">" & vbcrlf
   response.write "  <caption align=""left""><b>Attendance Per Hour</b></caption>" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <th width=""80%"">Time</th>" & vbcrlf
   response.write "      <th align=""center"">Total</th>" & vbcrlf
   response.write "  </tr>" & vbcrlf

  'Retrieve all of the dates
   sSQL = "SELECT distinct pool_date "
   sSQL = sSQL & " FROM egov_pool_info_vw "
   sSQL = SSQL & " WHERE orgid = " & session("orgid")
   sSQL = SSQL & " AND poolinfoid in (" & p_poolinfo_ids & ")"
   sSQL = sSQL & " ORDER BY pool_date "

   set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open sSQL, Application("DSN"), 3, 1

   if not rs.eof then
      lcl_dates = ""
      while not rs.eof
         if lcl_dates <> "" then
            lcl_dates = lcl_dates & ",'" & CDATE(rs("pool_date")) & "'"
         else
            lcl_dates = "'" & CDATE(rs("pool_date")) & "'"
         end if

         rs.movenext
      wend
   end if

  'Cycle through the date(s) to get a distinct list of hours
   if lcl_dates <> "" then
      sSQLh = "SELECT DISTINCT DATEPART(HH,scan_datetime) AS scan_hour "
      sSQLh = sSQLh & " FROM egov_pool_attendance_log "
      sSQLh = sSQLh & " WHERE orgid = " & session("orgid")
      sSQLh = sSQLh & " AND CAST(CONVERT(varchar(10), scan_datetime, 101) AS datetime) IN (" & lcl_dates & ")"
      sSQLh = sSQLh & " ORDER BY 1 "

      set oHours = Server.CreateObject("ADODB.Recordset")
      oHours.Open sSQLh, Application("DSN"), 3, 1

      if not oHours.eof then
         while not oHours.eof
           'Display the hour values properly (convert from 24 hour to 12 hour values)
            if oHours("scan_hour") > 12 then
               lcl_scan_hour = oHours("scan_hour") - 12
            else
               lcl_scan_hour = oHours("scan_hour")
            end if

            if oHours("scan_hour") < 12 then
               lcl_ampm = "AM"
            else
               lcl_ampm = "PM"
            end if

            if lcl_scan_hour < 1 then
               lcl_scan_hour = 12
            end if

'            if lcl_scan_hour < 10 then
'               lcl_scan_hour = "0" & lcl_scan_hour
'            end if

            response.write "  <tr>" & vbcrlf
            response.write "      <td>" & lcl_scan_hour & ":00 " & lcl_ampm & "</td>" & vbcrlf


           'Sum the member count per hour
            sSQL1 = "SELECT distinct memberid, people_count "
            sSQL1 = sSQL1 & " FROM egov_pool_attendance_log "
            sSQL1 = sSQL1 & " WHERE orgid = " & session("orgid")
            sSQL1 = sSQL1 & " AND CAST(CONVERT(varchar(10), scan_datetime, 101) AS datetime) IN (" & lcl_dates & ")"
            sSQL1 = sSQL1 & " AND DATEPART(HH,scan_datetime) = " & oHours("scan_hour")
            sSQL1 = sSQL1 & " AND memberid IS NOT NULL "
            sSQL1 = sSQL1 & " AND memberid <> '' "

            set oMemberCount = Server.CreateObject("ADODB.Recordset")
            oMemberCount.Open sSQL1, Application("DSN"), 3, 1

            lcl_member_count = 0

            if not oMemberCount.eof then
               while not oMemberCount.eof
                  lcl_member_count = lcl_member_count + oMemberCount("people_count")

                  oMemberCount.movenext
               wend
            end if

           'Sum the punchards, guests, and groups per hour
            sSQL2 = "SELECT distinct attendancelogid, people_count "
            sSQL2 = sSQL2 & " FROM egov_pool_attendance_log "
            sSQL2 = sSQL2 & " WHERE orgid = " & session("orgid")
            sSQL2 = sSQL2 & " AND CAST(CONVERT(varchar(10), scan_datetime, 101) AS datetime) IN (" & lcl_dates & ")"
            sSQL2 = sSQL2 & " AND DATEPART(HH,scan_datetime) = " & oHours("scan_hour")
            sSQL2 = sSQL2 & " AND (memberid IS NULL OR memberid = '') "

            set oCount = Server.CreateObject("ADODB.Recordset")
            oCount.Open sSQL2, Application("DSN"), 3, 1

            lcl_custom_count = 0

            if not oCount.eof then
               while not oCount.eof
                  lcl_custom_count = lcl_custom_count + oCount("people_count")

                  oCount.movenext
               wend
            end if

            lcl_total = lcl_member_count + lcl_custom_count

            response.write "      <td align=""center"">" & lcl_total & "</td>" & vbcrlf
            response.write "  </tr>" & vbcrlf

            set oMemberCount = nothing
            set oCount       = nothing

            oHours.movenext
         wend
      else
         response.write "  <tr>" & vbcrlf
         response.write "      <td colspan=""2"">No Records Exist</td>" & vbcrlf
         response.write "  </tr>" & vbcrlf
      end if
   end if

   response.write "</table>" & vbcrlf
   response.write "</div>" & vbcrlf
end sub

'--------------------------------------------------------------------------
 sub displayWeatherAverages(p_poolinfo_ids)
   lcl_temp_air_label   = "Air<br>Temperature"
   lcl_temp_water_label = "Water<br>Temperature"

   if instr(p_poolinfo_ids,",") > 0 then
      lcl_temp_air_label   = "Avg. " & lcl_temp_air_label
      lcl_temp_water_label = "Avg. " & lcl_temp_water_label
   end if

   response.write "<p>" & vbcrlf
   response.write "<div align=""left"">" & vbcrlf
	 	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""width: 700px"">" & vbcrlf
   response.write "  <caption align=""left""><b>Weather</b></caption>" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <th width=""15%"">Time</th>" & vbcrlf
   response.write "      <th width=""10%"" align=""center"">" & lcl_temp_air_label   & "</th>" & vbcrlf
   response.write "      <th width=""10%"" align=""center"">" & lcl_temp_water_label & "</th>" & vbcrlf

   if instr(p_poolinfo_ids,",") = 0 then
      response.write "      <th width=""65%"">Description</th>" & vbcrlf
   end if

   response.write "  </tr>" & vbcrlf

  'Retrieve all of the totals
   if instr(p_poolinfo_ids,",") > 0 then
     'If viewing the SUMMARY then we need to get a distinct list of weather_times
      sSQL = "SELECT DISTINCT weather_time, RIGHT(weather_time,2), LEFT(weather_time,2) "
      sSQL = sSQL & " FROM egov_pool_weather_log "
      sSQL = sSQL & " WHERE poolinfoid IN (" & p_poolinfo_ids & ")"
      sSQL = sSQL & " ORDER BY RIGHT(weather_time,2), LEFT(weather_time,2) "
   else
     'If viewing the DAILY report then simply retrieve the records
      sSQL = "SELECT weather_time, temperature_air, temperature_water, description "
      sSQL = sSQL & " FROM egov_pool_weather_log "
      sSQL = sSQL & " WHERE poolinfoid = " & p_poolinfo_ids
      sSQL = sSQL & " ORDER BY RIGHT(weather_time,2), LEFT(weather_time,2) "
   end if

   set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open sSQL, Application("DSN"), 3, 1

   if not rs.eof then
      lcl_bgcolor = "#eeeeee"
      while not rs.eof
         lcl_bgcolor           = changeBGColor(lcl_bgcolor,"","")
         lcl_temperature_air   = "&nbsp;"
         lcl_temperature_water = "&nbsp;"

        'If viewing the SUMMARY then we need to get averages for the Air/Water Temperatures
         if instr(p_poolinfo_ids,",") > 0 then
            sSQLa = "SELECT AVG(temperature_air) as temperature_air, AVG(temperature_water) as temperature_water "
            sSQLa = sSQLa & " FROM egov_pool_weather_log "
            sSQLa = sSQLa & " WHERE poolinfoid IN (" & p_poolinfo_ids & ")"
            sSQLa = sSQLa & " AND weather_time = '" & rs("weather_time") & "' "

            set rsa = Server.CreateObject("ADODB.Recordset")
            rsa.Open sSQLa, Application("DSN"), 3, 1

            if not rsa.eof then
               lcl_temperature_air   = rsa("temperature_air")
               lcl_temperature_water = rsa("temperature_water")
            end if

            set rsa = nothing
         else
           'If viewing the DAILY report then simply pull the records
            lcl_temperature_air   = rs("temperature_air")
            lcl_temperature_water = rs("temperature_water")
         end if

         response.write "  <tr bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
         response.write "      <td>" & rs("weather_time") & "</td>" & vbcrlf
         response.write "      <td align=""center"">" & lcl_temperature_air   & "</td>" & vbcrlf
         response.write "      <td align=""center"">" & lcl_temperature_water & "</td>" & vbcrlf

        'If viewing the SUMMARY then do NOT display the description
         if instr(p_poolinfo_ids,",") = 0 then
            response.write "      <td>" & rs("description") & "</td>" & vbcrlf
         end if

         response.write "  </tr>" & vbcrlf

         rs.movenext
      wend

      set rs = nothing
   else
      response.write "  <tr>" & vbcrlf
      response.write "      <td colspan=""4"">No Records Exist</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
   end if

   response.write "</table>" & vbcrlf
   response.write "</div>" & vbcrlf
 end sub

'-----------------------------------------------------------------------
 sub displayIncidents(p_poolinfo_ids)
   response.write "<p>" & vbcrlf
   response.write "<div align=""left"">" & vbcrlf
	 	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""width: 700px"">" & vbcrlf
   response.write "  <caption align=""left""><b>Incidents</b></caption>" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <th>Date</th>" & vbcrlf
   response.write "      <th width=""80%"">Incident</th>" & vbcrlf
   response.write "  </tr>" & vbcrlf

  'Retrieve all of the totals
   sSQLi = "SELECT i.incidentid, i.poolinfoid, i.incident_time, i.incident_time_ampm, i.name_of_injured, i.injury_type, i.witness, "
   sSQLi = sSQLi & " i.staff_response, i.report_completed_by, i.report_completed_by_datetime, p.pool_date "
   sSQLi = sSQLi & " FROM egov_pool_incidents_log i, egov_pool_info p "
   sSQLi = sSQLi & " WHERE i.poolinfoid = p.poolinfoid "
   sSQLi = sSQLi & " AND i.orgid = " & session("orgid")
   sSQLi = SSQLi & " AND i.poolinfoid in (" & p_poolinfo_ids & ")"
   sSQLi = sSQLi & " ORDER BY p.pool_date, i.incident_time_ampm, CAST(REPLACE(LEFT(i.incident_time,2),':','') AS INT), CAST(REPLACE(RIGHT(i.incident_time,2),':','') AS INT) "

   set rsi = Server.CreateObject("ADODB.Recordset")
   rsi.Open sSQLi, Application("DSN"), 3, 1

   if not rsi.eof then
      lcl_bgcolor = "#eeeeee"
      while not rsi.eof
         lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")

        'Set up the incident date and/or time
         lcl_incident_date = rsi("incident_time") & " " & rsi("incident_time_ampm")

         if instr(p_poolinfo_ids,",") > 0 then
            lcl_incident_date = rsi("pool_date") & " " & lcl_incident_date
         end if

         response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
         response.write "      <td width=""20%"">" & lcl_incident_date & "</td>" & vbcrlf
         response.write "      <td width=""80%"" style=""padding-top: 4px"">" & vbcrlf
         response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"" width=""100%"" style=""background-color: " & lcl_bgcolor & """>" & vbcrlf
         response.write "            <tr>" & vbcrlf
         response.write "                <td width=""30%""><b>Name of Injured:</b></td>" & vbcrlf
         response.write "                <td width=""70%"">" & rsi("name_of_injured") & "</td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "            <tr>" & vbcrlf
         response.write "                <td><b>Type of Injury:</b></td>" & vbcrlf
         response.write "                <td>" & rsi("injury_type") & "</td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "            <tr>" & vbcrlf
         response.write "                <td><b>Witness:</b></td>" & vbcrlf
         response.write "                <td>" & rsi("witness") & "</td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "            <tr valign=""top"">" & vbcrlf
         response.write "                <td><b>Staff Response:</b></td>" & vbcrlf
         response.write "                <td>" & rsi("staff_response") & "</td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "            <tr>" & vbcrlf
         response.write "                <td><b>Report Completed By:</b></td>" & vbcrlf
         response.write "                <td>" & rsi("report_completed_by") & "</td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "            <tr>" & vbcrlf
         response.write "                <td><b>Report Completed Date:</b></td>" & vbcrlf
         response.write "                <td>" & rsi("report_completed_by_datetime") & "</td>" & vbcrlf
         response.write "            </tr>" & vbcrlf
         response.write "          </table>" & vbcrlf
         response.write "      </td>" & vbcrlf
         response.write "  </tr>" & vbcrlf

         rsi.movenext
      wend

      set rsi = nothing

   else
      response.write "  <tr>" & vbcrlf
      response.write "      <td colspan=""2"">No Records Exist</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
   end if

   response.write "</table>" & vbcrlf
   response.write "</div>" & vbcrlf
 end sub

'-----------------------------------------------------------------------
 sub displayNotes(p_poolinfo_ids)
   response.write "<p>" & vbcrlf
   response.write "<div align=""left"">" & vbcrlf
	 	response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" class=""tableadmin"" style=""width: 700px"">" & vbcrlf
   response.write "  <caption align=""left""><b>Other Notes</b></caption>" & vbcrlf
   response.write "  <tr>" & vbcrlf
   response.write "      <th>Submitted Date</th>" & vbcrlf
   response.write "      <th>Submitted By</th>" & vbcrlf
   response.write "      <th width=""60%"">Note</th>" & vbcrlf
   response.write "  </tr>" & vbcrlf

  'Retrieve all of the totals
   sSQLn = "SELECT noteid, note_submittedby, note_datetime, description "
   sSQLn = sSQLn & " FROM egov_pool_info_notes "
   sSQLn = sSQLn & " WHERE poolinfoid IN (" & p_poolinfo_ids & ") "
   sSQLn = sSQLn & " AND orgid = " & session("orgid")
   sSQLn = sSQLn & " ORDER BY note_datetime DESC "

   set rsn = Server.CreateObject("ADODB.Recordset")
   rsn.Open sSQLn, Application("DSN") , 3, 1

   if not rsn.eof then
      lcl_bgcolor = "#eeeeee"
      while not rsn.eof
         lcl_bgcolor = changeBGColor(lcl_bgcolor,"","")

         response.write "  <tr valign=""top"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
         response.write "      <td>" & rsn("note_datetime")    & "</td>" & vbcrlf
         response.write "      <td>" & rsn("note_submittedby") & "</td>" & vbcrlf
         response.write "      <td>" & rsn("description")      & "</td>" & vbcrlf
         response.write "  </tr>" & vbcrlf

         rsn.movenext
      wend

      set rsn = nothing

   else
      response.write "  <tr>" & vbcrlf
      response.write "      <td colspan=""3"">No Records Exist</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
   end if

   response.write "</table>" & vbcrlf
   response.write "</div>" & vbcrlf
 end sub
%>