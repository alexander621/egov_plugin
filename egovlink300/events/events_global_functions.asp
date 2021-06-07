<%
'------------------------------------------------------------------------------
function getFeatureByID(p_orgid, p_featureid)
 lcl_return = ""

	'sSQL = "SELECT ISNULL(FO.featurename,F.featurename) AS featurename "
	sSQL = "SELECT f.feature "
	sSQL = sSQL & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSQL = sSQL & " WHERE FO.featureid = F.featureid "
 sSQL = sSQL & " AND FO.orgid = " & p_orgid
 sSQL = sSQL & " AND f.featureid = " & p_featureid

	set oGetFeatureByID = Server.CreateObject("ADODB.Recordset")
	oGetFeatureByID.Open sSQL, Application("DSN"), 0, 1

	if not oGetFeatureByID.eof then
  		lcl_return = oGetFeatureByID("feature")
 end if

	oGetFeatureByID.Close
	set oRs = nothing

 getFeatureByID = lcl_return

end function

'------------------------------------------------------------------------------
sub getEventCategoryOptions(p_orgid, p_calendarfeature, p_selected)

   'Get and display the options, if any exist, in the Event Category dropdown list.
    sSQL = "SELECT * "
    sSQL = sSQL & " FROM eventcategories "
    sSQL = sSQL & " WHERE orgid = " & p_orgid

    if p_calendarfeature <> "" then
       sSQL = sSQL & " AND UPPER(calendarfeature) = '" & UCASE(dbsafe(p_calendarfeature)) & "'"
    else
       sSQL = sSQL & " AND (calendarfeature = '' OR calendarfeature IS NULL)"
    end if

    sSQL = sSQL & " ORDER BY categoryname "

    set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sSQL, Application("DSN"), 3, 1

    if not rs.eof then
       while NOT rs.eof
         'Determine if the color is selected
          if rs("CategoryID") = p_selected then
             sSel = " selected=""selected"""
   					  else
			   		     sSel = ""
          end if

         'Set the class color of the option
          lcl_classcolor = ""

          if rs("color") <> "" then
             lcl_classcolor = " class=""color" & ucase(replace(rs("color"),"#","")) & """"
          end if

          response.write "<option value=""" & rs("CategoryID") & """" & lcl_classcolor & sSel & ">" & rs("CategoryName") & "</option>" & vbcrlf

          rs.movenext
       wend
    end if

    set rs = nothing
end sub

'---------------------------------------------------------------
 function getCategoryName(p_categoryid)
   lcl_return = ""

   sSQL = "SELECT categoryname "
   sSQL = sSQL & " FROM eventcategories "
   sSQL = sSQL & " WHERE categoryid = " & p_categoryid

   set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open sSQL, Application("DSN"), 3, 1

   if not rs.eof then
      lcl_return = rs("categoryname")
   end if

   set rs = nothing

   getCategoryName = lcl_return

 end function

'---------------------------------------------------------------
 sub displayCustomCalendarOptions(p_orgid, p_calendarfeatureid)
		 'sSQL = "SELECT isnull(FO.featurename,F.featurename) as featurename, F.feature "
		 sSQL = "SELECT F.featureid, isnull(FO.featurename,F.featurename) as featurename, F.feature "
		 sSQL = sSQL & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		 sSQL = sSQL & "WHERE "
   'sSQL = sSQL & " FO.publiccanview = 1 "
   sSQL = sSQL & " F.haspublicview = 1 "
   sSQL = sSQL & " AND O.orgid = FO.orgid "
   sSQL = sSQL & " AND FO.featureid = F.featureid "
   sSQL = sSQL & " AND O.orgid = " & p_orgid
   sSQL = sSQL & " AND f.parentfeatureid = (select f2.featureid from egov_organization_features as f2 where UPPER(f2.feature) = 'CUSTOM_CALENDARS')"
		 sSQL = sSQL & " ORDER BY FO.publicdisplayorder,F.publicdisplayorder"

		 Set oExists = Server.CreateObject("ADODB.Recordset")
		 oExists.Open sSQL, Application("DSN"), 3, 1

   if NOT oExists.eof then
      while NOT oExists.eof
         'if UCASE(oExists("feature")) = UCASE(p_calendarfeature) then
         if oExists("featureid") = p_calendarfeatureid then
            lcl_selected = " selected=""selected"""
         else
            lcl_selected = ""
         end if

         'response.write "  <option value=""" & oExists("feature") & """" & lcl_selected & ">" & oExists("featurename") & "</option>" & vbcrlf
         response.write "  <option value=""" & oExists("featureid") & """" & lcl_selected & ">" & oExists("featurename") & "</option>" & vbcrlf

         oExists.movenext
      wend
   end if

   set oExists = nothing

 end sub

'------------------------------------------------------------------------------
 function checkForCustomCalendars(p_orgid)
   lcl_return = "N"

		 sSQL = "SELECT distinct 'Y' as lcl_exists "
		 sSQL = sSQL & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		 sSQL = sSQL & "WHERE "
'   sSQL = sSQL & " FO.publiccanview = 1 "
   sSQL = sSQL & " F.haspublicview = 1 "
   sSQL = sSQL & " AND O.orgid = FO.orgid "
   sSQL = sSQL & " AND FO.featureid = F.featureid "
   sSQL = sSQL & " AND O.orgid = " & p_orgid
   sSQL = sSQL & " AND f.parentfeatureid = (select f2.featureid from egov_organization_features as f2 where UPPER(f2.feature) = 'CUSTOM_CALENDARS')"

		 Set oExists = Server.CreateObject("ADODB.Recordset")
		 oExists.Open sSQL, Application("DSN"), 3, 1

   if NOT oExists.eof then
      lcl_return = oExists("lcl_exists")
   end if

   checkForCustomCalendars = lcl_return

   set oExists = nothing

 end function

'------------------------------------------------------------------------------
function getItemsPerDay(p_orgid)
  lcl_return = 4

  if p_orgid <> "" then
     sSQL = "SELECT isnull(calendar_numItemsPerDay,4) as calendar_numItemsPerDay "
     sSQL = sSQL & " FROM organizations "
     sSQL = sSQL & " WHERE orgid = " & p_orgid

  		 set oGetItemsPerDay = Server.CreateObject("ADODB.Recordset")
		   oGetItemsPerDay.Open sSQL, Application("DSN"), 3, 1

     if not oGetItemsPerDay.eof then
        lcl_return = oGetItemsPerDay("calendar_numItemsPerDay")
     end if

     oGetItemsPerDay.close
     set oGetItemsPerDay = nothing

  end if

  getItemsPerDay = lcl_return

end function

'------------------------------------------------------------------------------
sub displayHistoryInfo(iDisplayHistoryOption, iFieldType, iUserID, iUserName, iHistoryDate)

  sDisplayHistoryOption = ""
  sFieldType            = "CREATEDBY"
  sFieldLabel           = "Created"
  sHistoryInfo          = ""

  if iDisplayHistoryOption <> "" then
     sDisplayHistoryOption = ucase(iDisplayHistoryOption)
  end if

  if iFieldType <> "" then
     sFieldType = ucase(iFieldType)
  end if

  if sFieldType = "LASTUPDATED" then
     sFieldLabel = "Last Updated"
  end if

  if iUserID <> "" then
     if sDisplayHistoryOption = "NAMES ONLY" then
        lcl_display_label = sFieldLabel & " By"
        sHistoryInfo = iUserName
     elseif sDisplayHistoryOption = "DATE/TIME ONLY" then
        lcl_display_label = sFieldLabel & " On"
        sHistoryInfo = iHistoryDate
     else
        lcl_display_label = sFieldLabel & " By"

        if iUserName <> "" then
           sHistoryInfo = iUserName
        end if

        if iHistoryDate <> "" then
           if sHistoryInfo <> "" then
              sHistoryInfo = sHistoryInfo & " on " & iHistoryDate
           else
              sHistoryInfo = iHistoryDate
           end if
        end if
     end if

     if sHistoryInfo <> "" then
        response.write "  <tr>" & vbcrlf
        response.write "      <td>" & lcl_display_label & ":</td>" & vbcrlf
        response.write "      <td style=""color:#800000"">" & sHistoryInfo & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if

  end if

end sub

'---------------------------------------------------------------
 function dbsafe(p_value)
    lcl_return = ""

    if trim(p_value) <> "" then
       lcl_return = replace(p_value,"'","''")
    end if

    dbsafe = lcl_return
 end function
%>
