<!-- #include file="../includes/common.asp" //-->
<%
'Retrieve all of the parameters
 lcl_poolinfoid   = request("poolinfoid")
 lcl_activetab_id = request("activetab")
 lcl_tabid        = request("tabid")

'Check to see if this record is to be deleted.
if request("cmd") = "D" then
   sSQLd = ""

   if UCASE(lcl_tabid) = "WEATHER" then

      lcl_weather_id = request("weatherid")
      sSQLd = "DELETE FROM egov_pool_weather_log WHERE weatherid = " & lcl_weather_id

   elseif UCASE(lcl_tabid) = "INCIDENT" then

      lcl_incident_id = request("incidentid")
      sSQLd = "DELETE FROM egov_pool_incidents_log WHERE incidentid = " & lcl_incident_id

   elseif UCASE(lcl_tabid) = "NOTE" then

      lcl_note_id = request("noteid")
      sSQLd = "DELETE FROM egov_pool_info_notes WHERE noteid = " & lcl_note_id

   end if

  'Delete the weather record
   if sSQLd <> "" then
     	set rsd = Server.CreateObject("ADODB.Recordset")
  	   rsd.Open sSQLd, Application("DSN") , 3, 1
   end if

'---------------------------------------------------------------------
else
'---------------------------------------------------------------------
 if UCASE(request("tabid")) = "WEATHER_ADD" then
'---------------------------------------------------------------------
    lcl_weather_time      = "NULL"
    lcl_temperature_air   = 0
    lcl_temperature_water = 0
    lcl_description       = "NULL"

    if request("p_new_weather_time") <> "" then
       lcl_weather_time = "'" & dbsafe(request("p_new_weather_time")) & "'"
    end if

    if request("p_new_weather_temp_air") <> "" then
       lcl_temperature_air   = request("p_new_weather_temp_air")
    end if

    if request("p_new_weather_temp_water") <> "" then
       lcl_temperature_water = request("p_new_weather_temp_water")
    end if

    if request("p_new_weather_description") <> "" then
       lcl_description = "'" & dbsafe(request("p_new_weather_description")) & "'"
    end if

    sSQLwa = "INSERT INTO egov_pool_weather_log ("
    sSQLwa = sSQLwa & "poolinfoid, "
    sSQLwa = sSQLwa & "orgid, "
    sSQLwa = sSQLwa & "weather_time, "
    sSQLwa = sSQLwa & "temperature_air, "
    sSQLwa = sSQLwa & "temperature_water, "
    sSQLwa = sSQLwa & "description"
    sSQLwa = sSQLwa & ") VALUES ("
    sSQLwa = sSQLwa & lcl_poolinfoid        & ", "
    sSQLwa = sSQLwa & session("orgid")      & ", "
    sSQLwa = sSQLwa & lcl_weather_time      & ", "
    sSQLwa = sSQLwa & lcl_temperature_air   & ", "
    sSQLwa = sSQLwa & lcl_temperature_water & ", "
    sSQLwa = sSQLwa & lcl_description
    sSQLwa = sSQLwa & ")"

   	set rswa = Server.CreateObject("ADODB.Recordset")
   	rswa.Open sSQLwa, Application("DSN") , 3, 1

    lcl_success = "SA"

'---------------------------------------------------------------------
 elseif UCASE(request("tabid")) = "INCIDENT_ADD" then
'---------------------------------------------------------------------
    lcl_incidenttime        = "NULL"
    lcl_incidenttime_hour   = "NULL"
    lcl_incidenttime_minute = "NULL"
    lcl_incidenttime_ampm   = "'" & request("p_new_incident_time_ampm") & "'"
    lcl_nameofinjured       = "NULL"
    lcl_typeofinjury        = "NULL"
    lcl_witness             = "NULL"
    lcl_staffresponse       = "NULL"
    lcl_completedby         = "NULL"

    if request("p_new_incident_time_hour") <> "" then
       lcl_incidenttime_hour = dbsafe(request("p_new_incident_time_hour"))
    end if

    if request("p_new_incident_time_min") <> "" then
       lcl_incidenttime_minute = dbsafe(request("p_new_incident_time_min"))
    end if

    lcl_incidenttime = "'" & lcl_incidenttime_hour & ":" & lcl_incidenttime_minute & "'"

    if request("p_new_incident_nameofinjured") <> "" then
       lcl_nameofinjured = "'" & dbsafe(request("p_new_incident_nameofinjured")) & "'"
    end if

    if request("p_new_incident_typeofinjury") <> "" then
       lcl_typeofinjury = "'" & dbsafe(request("p_new_incident_typeofinjury")) & "'"
    end if

    if request("p_new_incident_witness") <> "" then
       lcl_witness = "'" & dbsafe(request("p_new_incident_witness")) & "'"
    end if

    if request("p_new_incident_staffresponse") <> "" then
       lcl_staffresponse = "'" & dbsafe(request("p_new_incident_staffresponse")) & "'"
    end if

    if request("p_new_incident_completedby") <> "" then
       lcl_completedby = "'" & dbsafe(request("p_new_incident_completedby")) & "'"
    end if

    if request("p_new_incident_staffresponse") <> "" then
       lcl_staffresponse = "'" & dbsafe(request("p_new_incident_staffresponse")) & "'"
    end if

    sSQLia = "INSERT INTO egov_pool_incidents_log ("
    sSQLia = sSQLia & "poolinfoid, "
    sSQLia = sSQLia & "orgid, "
    sSQLia = sSQLia & "incident_time, "
    sSQLia = sSQLia & "incident_time_ampm, "
    sSQLia = sSQLia & "name_of_injured, "
    sSQLia = sSQLia & "injury_type, "
    sSQLia = sSQLia & "witness, "
    sSQLia = sSQLia & "staff_response, "
    sSQLia = sSQLia & "report_completed_by, "
    sSQLia = sSQLia & "report_completed_by_datetime"
    sSQLia = sSQLia & ") VALUES ("
    sSQLia = sSQLia & lcl_poolinfoid        & ", "
    sSQLia = sSQLia & session("orgid")      & ", "
    sSQLia = sSQLia & lcl_incidenttime      & ", "
    sSQLia = sSQLia & lcl_incidenttime_ampm & ", "
    sSQLia = sSQLia & lcl_nameofinjured     & ", "
    sSQLia = sSQLia & lcl_typeofinjury      & ", "
    sSQLia = sSQLia & lcl_witness           & ", "
    sSQLia = sSQLia & lcl_staffresponse     & ", "
    sSQLia = sSQLia & lcl_completedby       & ", "
    sSQLia = sSQLia & "'" & Now()           & "'"
    sSQLia = sSQLia & ")"

   	set rsia = Server.CreateObject("ADODB.Recordset")
   	rsia.Open sSQLia, Application("DSN") , 3, 1

    lcl_success = "SA"

'---------------------------------------------------------------------
 elseif UCASE(request("tabid")) = "NOTES_ADD" then
'---------------------------------------------------------------------
    lcl_submittedby = "NULL"
    lcl_note        = "NULL"

    if request("p_new_submittedby") <> "" then
       lcl_submittedby = "'" & dbsafe(request("p_new_submittedby")) & "'"
    end if

    if request("p_new_note") <> "" then
       lcl_note = "'" & dbsafe(request("p_new_note")) & "'"
    end if

    sSQLna = "INSERT INTO egov_pool_info_notes ("
    sSQLna = sSQLna & "poolinfoid, "
    sSQLna = sSQLna & "orgid, "
    sSQLna = sSQLna & "note_submittedby, "
    sSQLna = sSQLna & "note_datetime, "
    sSQLna = sSQLna & "description "
    sSQLna = sSQLna & ") VALUES ("
    sSQLna = sSQLna & lcl_poolinfoid   & ", "
    sSQLna = sSQLna & session("orgid") & ", "
    sSQLna = sSQLna & lcl_submittedby  & ", "
    sSQLna = sSQLna & "'" & Now()      & "', "
    sSQLna = sSQLna & lcl_note
    sSQLna = sSQLna & ")"

   	set rsna = Server.CreateObject("ADODB.Recordset")
   	rsna.Open sSQLna, Application("DSN") , 3, 1

    lcl_success = "SA"

'---------------------------------------------------------------------
  else
' elseif UCASE(request("tabid")) = "SAVE" or request("tabid") = "" then
'---------------------------------------------------------------------

    lcl_pool_date = request("pool_date")

    if lcl_poolinfoid = "" then
       lcl_poolinfoid = 0
    end if

    if CLng(lcl_poolinfoid) = CLng(0) then
      'Check to see if a record already exists with the same date.
       if checkForPoolDate(lcl_pool_date) then
          response.redirect "pool_attendance_maint.asp?pid=0&success=AE&pool_date=" & lcl_pool_date
       end if

      'Set up the parameters that are to be inserted.
       sSQL = "INSERT INTO egov_pool_info ("
       sSQL = sSQL & "orgid, "
       sSQL = sSQL & "pool_date "

      'set up the values to be inserted
       sSQL = sSQL & ") VALUES ("
       sSQL = sSQL       & session("orgid")      & ", "
       sSQL = sSQL & "'" & dbsafe(lcl_pool_date) & "'"
       sSQL = sSQL & ")"

       lcl_success = "SA"

    else
    		'Update existing record
		     sSQL = "UPDATE egov_pool_info SET "
       sSQL = sSQL & " pool_date = '" & dbsafe(lcl_pool_date) & "' "
		     sSQL = sSQL & " WHERE poolinfoid = " & lcl_poolinfoid

       lcl_success = "SU"
    end if

   	set rs = Server.CreateObject("ADODB.Recordset")
   	rs.Open sSQL, Application("DSN") , 3, 1

   'Retrieve the poolinfoid that was just inserted
    if lcl_poolinfoid = 0 then
       sSQLid = "SELECT IDENT_CURRENT('egov_pool_info') as NewID"
       rs.Open sSQLid, Application("DSN") , 3, 1
       lcl_identity = rs.Fields("NewID").value

      	Set oCmd = Nothing

       if lcl_identity <> "" AND (lcl_poolinfoid = 0 OR isnull(lcl_poolinfoid)) then
          lcl_poolinfoid = lcl_identity
       end if
    else
      'Update the remaining tabs

      'WEATHER_EDIT  -----------------------------------------------------------------
       for e = 1 to request("p_total_weather")				
          if request("p_weatherid_" & e)           <> "" OR _
             request("p_weather_time_" & e)        <> "" OR _
             request("p_temperature_air_" & e)     <> "" OR _
             request("p_temperature_water_" & e)   <> "" OR _
             request("p_weather_description_" & e) <> "" then
'             request("p_weather_delete_" & e)      <> "" then

             lcl_weatherid           = 0
             lcl_weather_time        = "NULL"
             lcl_temperature_air     = 0
             lcl_temperature_water   = 0
             lcl_weather_description = "NULL"
'             lcl_weather_delete      = "N"

             if request("p_weatherid_" & e) <> "" then
                lcl_weatherid = request("p_weatherid_" & e)
             end if

             if request("p_weather_time_" & e) <> "" then
                lcl_weather_time = "'" & dbsafe(request("p_weather_time_" & e)) & "'"
             end if

             if request("p_temperature_air_" & e) <> "" then
                lcl_temperature_air = request("p_temperature_air_" & e)
             end if

             if request("p_temperature_water_" & e) <> "" then
                lcl_temperature_water = request("p_temperature_water_" & e)
             end if

             if request("p_weather_description_" & e) <> "" then
                lcl_weather_description = "'" & dbsafe(request("p_weather_description_" & e)) & "'"
             end if

'             if request("p_weather_delete_" & e) <> "" then
'                lcl_weather_delete = request("p_weather_delete_" & e)
'             end if

'             if lcl_weather_delete = "N" then
            		  sSQLu = "UPDATE egov_pool_weather_log SET "
                sSQLu = sSQLu & " weather_time = "      & lcl_weather_time      & ", "
                sSQLu = sSQLu & " temperature_air = "   & lcl_temperature_air   & ", "
                sSQLu = sSQLu & " temperature_water = " & lcl_temperature_water & ", "
                sSQLu = sSQLu & " description = "       & lcl_weather_description
                sSQLu = sSQLu & " WHERE weatherid = " & lcl_weatherid

               	set rsu = Server.CreateObject("ADODB.Recordset")
            	   rsu.Open sSQLu, Application("DSN") , 3, 1
'             else
'            		  sSQLd = "DELETE FROM egov_pool_weather_log WHERE weatherid = " & lcl_weatherid

'               	set rsd = Server.CreateObject("ADODB.Recordset")
'            	   rsd.Open sSQLd, Application("DSN") , 3, 1
'             end if
          end if
       next

      'INCIDENT_EDIT  -----------------------------------------------------------------------
       for e = 1 to request("p_total_incidents")				
          if request("p_incidentid_" & e)          <> "" OR _
             request("p_incident_time_hour_" & e)  <> "" OR _
             request("p_incident_time_min_" & e)   <> "" OR _
             request("p_incident_time_ampm_" & e)  <> "" OR _
             request("p_name_of_injured_" & e)     <> "" OR _
             request("p_injury_type_" & e)         <> "" OR _
             request("p_witness_" & e)             <> "" OR _
             request("p_staff_response_" & e)      <> "" OR _
             request("p_report_completed_by_" & e) <> "" then
'             request("p_incident_delete_" & e)     <> "" then

             lcl_incidentid           = 0
             lcl_incidenttime         = "NULL"
             lcl_incidenttime_hour    = request("p_incident_time_hour_" & e)
             lcl_incidenttime_minute  = request("p_incident_time_min_"  & e)
             lcl_incidenttime_ampm    = "'" & request("p_incident_time_ampm_" & e) & "'"
             lcl_nameofinjured        = "NULL"
             lcl_typeofinjury         = "NULL"
             lcl_witness              = "NULL"
             lcl_staffresponse        = "NULL"
             lcl_completedby          = "NULL"
'             lcl_incident_delete      = "N"

             if request("p_incidentid_" & e) <> "" then
                lcl_incidentid = request("p_incidentid_" & e)
             end if

'             if request("p_incident_time_hour_" & e) <> "" then
'                lcl_incidenttime_hour = dbsafe(request("p_incident_time_hour_" & e))
'             end if

'             if request("p_incident_time_min_" & e) <> "" then
'                lcl_incidenttime_minute = dbsafe(request("p_incident_time_min_" & e))
'             end if

             lcl_incidenttime = "'" & lcl_incidenttime_hour & ":" & lcl_incidenttime_minute & "'"

'             if request("p_incident_time_ampm_" & e) <> "" then
'                lcl_incidenttime_ampm = "'" & dbsafe(request("p_incident_time_ampm_" & e)) & "'"
'             end if

             if request("p_name_of_injured_" & e) <> "" then
                lcl_nameofinjured = "'" & dbsafe(request("p_name_of_injured_" & e)) & "'"
             end if

             if request("p_injury_type_" & e) <> "" then
                lcl_typeofinjury = "'" & dbsafe(request("p_injury_type_" & e)) & "'"
             end if

             if request("p_witness_" & e) <> "" then
                lcl_witness = "'" & dbsafe(request("p_witness_" & e)) & "'"
             end if

             if request("p_staff_response_" & e) <> "" then
                lcl_staffresponse = "'" & dbsafe(request("p_staff_response_" & e)) & "'"
             end if

             if request("p_report_completed_by_" & e) <> "" then
                lcl_completedby = "'" & dbsafe(request("p_report_completed_by_" & e)) & "'"
             end if

'             if request("p_incident_delete_" & e) <> "" then
'                lcl_incident_delete = request("p_incident_delete_" & e)
'             end if

'             if lcl_incident_delete = "N" then
            		  sSQLu = "UPDATE egov_pool_incidents_log SET "
                sSQLu = sSQLu & " incident_time = "       & lcl_incidenttime      & ", "
                sSQLu = sSQLu & " incident_time_ampm = "  & lcl_incidenttime_ampm & ", "
                sSQLu = sSQLu & " name_of_injured = "     & lcl_nameofinjured     & ", "
                sSQLu = sSQLu & " injury_type = "         & lcl_typeofinjury      & ", "
                sSQLu = sSQLu & " witness = "             & lcl_witness           & ", "
                sSQLu = sSQLu & " staff_response = "      & lcl_staffresponse     & ", "
                sSQLu = sSQLu & " report_completed_by = " & lcl_completedby
                sSQLu = sSQLu & " WHERE incidentid = " & lcl_incidentid

               	set rsu = Server.CreateObject("ADODB.Recordset")
            	   rsu.Open sSQLu, Application("DSN") , 3, 1
'             else
'            		  sSQLd = "DELETE FROM egov_pool_incidents_log WHERE incidentid = " & lcl_incidentid

'               	set rsd = Server.CreateObject("ADODB.Recordset")
'         	      rsd.Open sSQLd, Application("DSN") , 3, 1
'             end if
          end if
       next

      'NOTES_EDIT  -----------------------------------------------------------------------------
       for e = 1 to request("p_total_notes")				
          if request("p_noteid_" & e)            <> "" OR _
             request("p_note_submittedby_" & e)  <> "" OR _
             request("p_notes_description_" & e) <> "" then
'             request("p_weather_delete_" & e)    <> "" then

             lcl_noteid            = 0
             lcl_note_submittedby  = "NULL"
             lcl_notes_description = "NULL"
'             lcl_notes_delete      = "N"

             if request("p_noteid_" & e) <> "" then
                lcl_noteid = request("p_noteid_" & e)
             end if

             if request("p_note_submittedby_" & e) <> "" then
                lcl_note_submittedby = "'" & dbsafe(request("p_note_submittedby_" & e)) & "'"
             end if

             if request("p_notes_description_" & e) <> "" then
                lcl_notes_description = "'" & dbsafe(request("p_notes_description_" & e)) & "'"
             end if

'             if request("p_notes_delete_" & e) <> "" then
'                lcl_notes_delete = request("p_notes_delete_" & e)
'             end if

'             if lcl_notes_delete = "N" then
            		  sSQLnu = "UPDATE egov_pool_info_notes SET "
                sSQLnu = sSQLnu & " note_submittedby = " & lcl_note_submittedby & ", "
                sSQLnu = sSQLnu & " description = "      & lcl_notes_description
                sSQLnu = sSQLnu & " WHERE noteid = " & lcl_noteid

               	set rsnu = Server.CreateObject("ADODB.Recordset")
            	   rsnu.Open sSQLnu, Application("DSN") , 3, 1
'             else
'            		  sSQLd = "DELETE FROM egov_pool_info_notes WHERE noteid = " & lcl_noteid

'               	set rsd = Server.CreateObject("ADODB.Recordset")
'            	   rsd.Open sSQLd, Application("DSN") , 3, 1
'             end if
          end if
       next

       lcl_success = "SU"

    end if

'---------------------------------------------------------------------
 end if

 lcl_enable_add  = "Y"
 lcl_enable_edit = "Y"

'Check the user roles to see when to redirect the user
 if NOT UserHasPermission( session("userid"), "pool_attendance_add" ) then
    lcl_enable_add = "N"
 end if

 if NOT UserHasPermission( session("userid"), "pool_attendance_edit" ) then
    lcl_enable_edit = "N"
 end if

 if lcl_enable_edit = "Y" then
   'Always ensure that returning to the maintenance screen sets the tab to SAVE
    lcl_tabid = "SAVE"

    response.redirect "pool_attendance_maint.asp?pid=" & lcl_poolinfoid & "&success=" & lcl_success & "&activetab=" & lcl_activetab_id & "&tabid=" & lcl_tabid
 else
    if lcl_enable_add = "Y" then
       response.redirect "pool_attendance_list.asp?success=" & lcl_success
    end if
 end if
'---------------------------------------------------------------------
end if  'end of check for delete
'---------------------------------------------------------------------
function dbsafe(p_value)
  if p_value <> "" then
     lcl_return = REPLACE(p_value,"'","''")
  else
     lcl_return = ""
  end if

  dbsafe = lcl_return

end function

'---------------------------------------------------------------------
function checkForPoolDate(p_pool_date)
  lcl_return = false

  sSQLc = "SELECT distinct 'Y' AS lcl_exists "
  sSQLc = sSQLc & " FROM egov_pool_info "
  sSQLc = sSQLc & " WHERE DATEDIFF(d,pool_date,'" & dbsafe(p_pool_date) & "') = 0 "
  sSQLc = sSQLc & " AND orgid = " & session("orgid")

  set oCheck = Server.CreateObject("ADODB.Recordset")
  oCheck.Open sSQLc, Application("DSN") , 3, 1

  if not oCheck.eof then
     lcl_return = true
  end if

  set oCheck = nothing

  checkForPoolDate = lcl_return

end function
%>