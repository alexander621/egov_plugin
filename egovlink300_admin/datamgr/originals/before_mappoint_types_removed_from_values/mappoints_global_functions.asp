<%
'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Successfully Updated..."
     elseif iSuccess = "SA" then
        lcl_return = "Successfully Created..."
     elseif iSuccess = "SR" then
        lcl_return = "Successfully Reordered..."
     elseif iSuccess = "SD" then
        lcl_return = "Successfully Deleted..."
     elseif iSuccess = "NE" then
        lcl_return = "Map-Point Category does not exist..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
function setupUserMaintLogInfo(iName, iDate)
  lcl_return = ""

  if iName <> "" then
     if lcl_return <> "" then
        lcl_return = lcl_return & iName
     else
        lcl_return = iName
     end if
  end if

  if iDate <> "" then
     if lcl_return <> "" then
        lcl_return = lcl_return & " on " & iDate
     else
        lcl_return = iDate
     end if
  end if

  setupUserMaintLogInfo = lcl_return

end function

'------------------------------------------------------------------------------
sub maintainMapPointTypeField(iMPFieldID, iMapPointTypeID, iOrgID, iFieldName, iFieldType, iDisplayInResults, iResultsOrder)

  if iFieldName <> "" then
     lcl_fieldname = "'" & dbsafe(iFieldName) & "'"
  else
     lcl_fieldname = "NULL"
  end if

  if iFieldType <> "" then
     lcl_fieldtype = "'" & dbsafe(UCASE(iFieldType)) & "'"
  else
     lcl_fieldtype = "NULL"
  end if

  if iDisplayInResults <> "" then
     lcl_displayInResults = iDisplayInResults
  else
     lcl_displayInResults = 0
  end if

  if iResultsOrder <> "" then
     lcl_resultsOrder = iResultsOrder
  else
     lcl_resultsOrder = 0
  end if

 'Determine if the Map-Point is to be added or updated
  if iMPFieldID <> "" then
     sSQL = "UPDATE egov_mappoints_types_fields SET "
     sSQL = sSQL & " fieldname = "        & lcl_fieldname        & ", "
     sSQL = sSQL & " fieldtype = "        & lcl_fieldtype        & ", "
     sSQL = sSQL & " displayInResults = " & lcl_displayInResults & ", "
     sSQL = sSQL & " resultsOrder = "     & lcl_resultsOrder
     sSQL = sSQL & " WHERE mp_fieldid = " & iMPFieldID

    'Check to see if the field exists on egov_mappoints_values.
    'If "yes" then update the "display" fields.
     'lcl_mpvalue_exists = checkMapPointValueExists(iMPFieldID)

     'if lcl_mpvalue_exists then
     '   sSQL2 = "UPDATE egov_mappoints_values SET "
     '   sSQL2 = sSQL2 & " fieldname = "        & lcl_fieldname        & ", "
     '   sSQL2 = sSQL2 & " fieldtype = "        & lcl_fieldtype        & ", "
     '   sSQL2 = sSQL2 & " displayInResults = " & lcl_displayInResults & ", "
     '   sSQL2 = sSQL2 & " resultsOrder = "     & lcl_resultsOrder
     '   sSQL2 = sSQL2 & " WHERE mp_fieldid = " & iMPFieldID
     'end if
  else

     sSQL = "INSERT INTO egov_mappoints_types_fields ("
     sSQL = sSQL & "mappoint_typeid,"
     sSQL = sSQL & "orgid,"
     sSQL = sSQL & "fieldname,"
     sSQL = sSQL & "fieldtype,"
     sSQL = sSQL & "displayInResults,"
     sSQL = sSQL & "resultsOrder"
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & iMapPointTypeID      & ", "
     sSQL = sSQL & iOrgID               & ", "
     sSQL = sSQL & lcl_fieldname        & ", "
     sSQL = sSQL & lcl_fieldtype        & ", "
     sSQL = sSQL & lcl_displayInResults & ", "
     sSQL = sSQL & lcl_resultsOrder
     sSQL = sSQL & ")"

  end if

 'Update/Insert the Map-Point Type Field
	 set oMaintainMapPointTypeField = Server.CreateObject("ADODB.Recordset")
'	 set oUpdateMapPointValue       = Server.CreateObject("ADODB.Recordset")

 'Maintain the Map-Point Type Field data on egov_mappoints_values
 	oMaintainMapPointTypeField.Open sSQL, Application("DSN"), 3, 1
' 	oUpdateMapPointValue.Open sSQL2, Application("DSN"), 3, 1

  set oMaintainMapPointTypeField = nothing
'  set oUpdateMapPointValue       = nothing

end sub

'------------------------------------------------------------------------------
sub deleteMapPointTypeField(iMPFieldID)

  sSQL = "DELETE FROM egov_mappoints_types_fields WHERE mp_fieldid = " & iMPFieldID

	 set oDeleteMapPointTypeField = Server.CreateObject("ADODB.Recordset")
 	oDeleteMapPointTypeField.Open sSQL, Application("DSN"), 3, 1

  set oDeleteMapPointTypeField = nothing

end sub

'------------------------------------------------------------------------------
sub deleteMapPointValue(iDBColumnName, iID)

  sSQL = "DELETE FROM egov_mappoints_values WHERE " & iDBColumnName & " = " & iID

	 set oDelMPValue = Server.CreateObject("ADODB.Recordset")
 	oDelMPValue.Open sSQL, Application("DSN"), 3, 1

  set oDelMPValue = nothing

end sub

'------------------------------------------------------------------------------
function checkMapPointValueExists(p_mp_fieldid)

  lcl_return = False
  lcl_exists = "N"

  if p_mp_fieldid <> "" then

     sSQL = "SELECT distinct 'Y' as mpvalue_exists "
     sSQL = sSQL & " FROM egov_mappoints_values "
     sSQL = sSQL & " WHERE mp_fieldid = " & p_mp_fieldid

   	 set oMPValueExists = Server.CreateObject("ADODB.Recordset")
    	oMPValueExists.Open sSQL, Application("DSN"), 3, 1

     if not oMPValueExists.eof then
        lcl_exists = oMPValueExists("mpvalue_exists")
     end if

     oMPValueExists.close
     set oMPValueExists = nothing

  end if

  if lcl_exists = "Y" then
     lcl_return = True
  end if

  checkMapPointValueExists = lcl_return

end function

'------------------------------------------------------------------------------
function RunIdentityInsert( sInsertStatement )
	 Dim sSQL, iReturnValue, oInsert

	 iReturnValue = 0

	'Insert new row into database and get rowid
 	sSQL = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

 	set oInsert = Server.CreateObject("ADODB.Recordset")
	 oInsert.Open sSQL, Application("DSN"), 3, 3

 	iReturnValue = oInsert("ROWID")

 	oInsert.close
	 set oInsert = nothing

 	RunIdentityInsert = iReturnValue

end function

'------------------------------------------------------------------------------
function checkForMapPointsByMapPointTypeID(iMapPointTypeID)
  lcl_return = True

  if iMapPointTypeID <> "" then
     sSQL = "SELECT DISTINCT 'Y' as lcl_exists "
     sSQL = sSQL & " FROM egov_mappoints "
     sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

    	set oCheckForMapPoints = Server.CreateObject("ADODB.Recordset")
   	 oCheckForMapPoints.Open sSQL, Application("DSN"), 3, 3

     if not oCheckForMapPoints.eof then
        lcl_return = False
     end if

     oCheckForMapPoints.close
     set oCheckForMapPoints = nothing

  end if

  checkForMapPointsByMapPointTypeID = lcl_return

end function

'------------------------------------------------------------------------------
function DisplayAddress(p_orgid, p_street_number, p_street_name)
	
	Dim sNumber, oAddressList, blnFound

 lcl_streetNumber    = p_street_number
 lcl_streetName      = p_street_name
 lcl_new_street_name = buildStreetAddress(p_street_number, "", p_street_name, "", "")

'Get list of addresses
	sSQL = "SELECT residentaddressid, isnull(residentstreetnumber,'') as residentstreetnumber, residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, streetsuffix, streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude "
 sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid=" & p_orgid
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND residentstreetname is not null "
 sSQL = sSQL & " ORDER BY sortstreetname, residentstreetprefix, Cast(residentstreetnumber as integer(4))"

	set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

 response.write "<input type=""hidden"" name=""residentstreetnumber"" id=""residentstreetnumber"" value="""" />" & vbcrlf
	response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkImportAddressBtn();"">" & vbcrlf
	response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf
	
	do while not oAddressList.eof
   'Build the original full street address
    lcl_original_street_name = buildStreetAddress(oAddressList("residentstreetnumber"), oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

   'Check for a matching address
    'if lcl_large_address_list then
    '   if trim(lcl_new_street_name) = trim(lcl_original_street_name) then
    '			    sSelected = " selected=""selected"""
    ' 		else
    '		    	sSelected = ""
    ' 		end If
    'else
    '   if p_street_number <> "" or not IsNull(p_street_number) then
    '      if trim(lcl_original_street_name) = trim(lcl_new_street_name) then
    '          sSelected = " selected=""selected"""
    '      else
    '          sSelected = ""
    '      end if
    '   else
          if lcl_streetName = oAddressList("residentstreetname") then
              sSelected = " selected=""selected"""
          else
              sSelected = ""
          end if
    '   end if
    'end if

   	response.write "  <option value=""" & oAddressList("residentaddressid") & """" & sSelected & ">" & lcl_original_street_name & "</option>" & vbcrlf

  		oAddressList.MoveNext
	loop

	response.write "</select>" & vbcrlf

	oAddressList.close
	set oAddressList = nothing

end function

'------------------------------------------------------------------------------
sub DisplayLargeAddressList(p_orgid, p_street_number, p_street_name)
 dim sSql, oAddressList

 lcl_streetNumber = p_street_number
 lcl_streetName   = p_street_name

	sSQL = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & p_orgid
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " ORDER BY sortstreetname "
	
 set oAddressList = Server.CreateObject("ADODB.Recordset")
 oAddressList.Open sSQL, Application("DSN"), 3, 1

 if not oAddressList.eof then
  		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetNumber & """ size=""8"" maxlength=""10"" onchange=""save_address();"" /> &nbsp; " & vbcrlf
 	 	response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkImportAddressBtn()"">" & vbcrlf
  		response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf

    do while not oAddressList.eof

      'Build the full street address
       sCompareName = buildStreetAddress("", oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

      'Determine if the option is selected
     		if lcl_streetName = sCompareName then
       			lcl_selected_address = " selected=""selected"""
       else
          lcl_selected_address = ""
     		end if

       response.write "<option value=""" & sCompareName & """" & lcl_selected_address & ">" & sCompareName & "</option>" & vbcrlf

    			oAddressList.MoveNext
    loop

    response.write "</select>&nbsp;" & vbcrlf
    response.write "<input type=""button"" id=""validateAddress""  class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
 end if

 oAddressList.close
 set oAddressList = nothing

end sub

'------------------------------------------------------------------------------
function formatFieldforInsertUpdate(p_value)
  lcl_return = "NULL"

  if trim(p_value) <> "" then
     lcl_return = "'" & dbsafe(p_value) & "'"
  end if

  formatFieldforInsertUpdate = lcl_return

end function

'------------------------------------------------------------------------------
function getMapPointTypeDescription(iMapPointTypeID)

  lcl_return = ""

  if iMapPointTypeID <> "" then
     lcl_mappointtypeid = CLng(iMapPointTypeID)
  else
     lcl_mappointtypeid = 0
  end if

  sSQL = "SELECT description "
  sSQL = sSQL & " FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE mappoint_typeid = " & lcl_mappointtypeid

  set oGetMPTDesc = Server.CreateObject("ADODB.Recordset")
  oGetMPTDesc.Open sSQL, Application("DSN"), 3, 1

  if not oGetMPTDesc.eof then
     lcl_return = oGetMPTDesc("description")
  end if

  oGetMPTDesc.close
  set oGetMPTDesc = nothing

  getMapPointTypeDescription = lcl_return

end function

'------------------------------------------------------------------------------
sub displayMapPointTypes(iOrgID, iMapPointTypeID)

  if iMapPointTypeID <> "" then
     lcl_mappointtypeid = CLng(iMapPointTypeID)
  else
     lcl_mappointtypeid = 0
  end if

  sSQL = "SELECT mappoint_typeid, description "
  sSQL = sSQL & " FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND isInactive = 0 "
  sSQL = sSQL & " ORDER BY description "

  set oDisplayMPTypes = Server.CreateObject("ADODB.Recordset")
  oDisplayMPTypes.Open sSQL, Application("DSN"), 3, 1

  if not oDisplayMPTypes.eof then
     do while not oDisplayMPTypes.eof

        if oDisplayMPTypes("mappoint_typeid") = lcl_mappointtypeid then
           lcl_selected_mappointtype = " selected=""selected"""
        else
           lcl_selected_mappointtype = ""
        end if

        response.write "  <option value=""" & oDisplayMPTypes("mappoint_typeid") & """" & lcl_selected_mappointtype & ">" & oDisplayMPTypes("description") & "</option>" & vbcrlf

        oDisplayMPTypes.movenext
     loop
  end if

end sub

'------------------------------------------------------------------------------
sub maintainMapPointValues(iOrgID, iMapPointTypeID, iMapPointID, iMPFieldID, iMPValueID, iFieldType, iFieldName, iFieldValue, iDisplayInResults, iResultsOrder)

  lcl_orgid           = iOrgID
  lcl_mappoint_typeid = iMapPointTypeID
  lcl_mappointid      = iMapPointID
  lcl_mp_fieldid      = iMPFieldID
  lcl_mp_valueid      = 0

  if iMPValueID <> "" then
     lcl_mp_valueid = iMPValueID
  end if

  if iDisplayInResults <> "" then
     if iDisplayInResults then
        lcl_displayinresults = 1
     else
        lcl_displayinresults = 0
     end if
  else
     lcl_displayinresults = 0
  end if

  if iResultsOrder <> "" then
     lcl_resultsorder = iResultsOrder
  else
     lcl_resultsorder = 1
  end if
    
  lcl_fieldtype  = formatFieldforInsertUpdate(iFieldType)
  lcl_fieldname  = formatFieldforInsertUpdate(iFieldName)
  lcl_fieldvalue = formatFieldforInsertUpdate(iFieldValue)

 'If a mp_valueid exists then update the Map-Point Value.  Otherwise, insert it.
  if lcl_mp_valueid > 0 then
     sSQL = "UPDATE egov_mappoints_values SET "
     sSQL = sSQL & " mappoint_typeid = "  & lcl_mappoint_typeid  & ", "
     sSQL = sSQL & " fieldname = "        & lcl_fieldname        & ", "
     sSQL = sSQL & " fieldvalue = "       & lcl_fieldvalue       & ", "
     sSQL = sSQL & " displayInResults = " & lcl_displayinresults
     sSQL = sSQL & " WHERE mp_valueid = " & lcl_mp_valueid

  else

     sSQL = "INSERT INTO egov_mappoints_values ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "mappoint_typeid, "
     sSQL = sSQL & "mappointid, "
     sSQL = sSQL & "mp_fieldid, "
     sSQL = sSQL & "fieldtype, "
     sSQL = sSQL & "fieldname, "
     sSQL = sSQL & "fieldvalue, "
     sSQL = sSQL & "displayInResults, "
     sSQL = sSQL & "resultsorder"
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & lcl_orgid            & ", "
     sSQL = sSQL & lcl_mappoint_typeid  & ", "
     sSQL = sSQL & lcl_mappointid       & ", "
     sSQL = sSQL & lcl_mp_fieldid       & ", "
     sSQL = sSQL & lcl_fieldtype        & ", "
     sSQL = sSQL & lcl_fieldname        & ", "
     sSQL = sSQL & lcl_fieldvalue       & ", "
     sSQL = sSQL & lcl_displayInResults & ", "
     sSQL = sSQL & "0"
     sSQL = sSQL & ")"

  end if

  set oMaintainMPValues = Server.CreateObject("ADODB.Recordset")
  oMaintainMPValues.Open sSQL, Application("DSN"), 3, 1

  set oMaintainMPValues = nothing

end sub

'------------------------------------------------------------------------------
function setupUrlParameters(iURLParameters, iFieldName, iFieldValue)
  lcl_return = ""

  if trim(iURLParameters) <> "" then
     lcl_return = iURLParameters
  end if

  if iFieldValue <> "" then
     if lcl_return <> "" then
        lcl_return = lcl_return & "&"
     else
        lcl_return = lcl_return & "?"
     end if

     lcl_return = lcl_return & iFieldName & "=" & iFieldValue

  end if

  setupUrlParameters = lcl_return

end function

'------------------------------------------------------------------------------
sub displayMPTypeStatuses(iOrgID, iStatusID, iLimitToActiveOnly)

  if iStatusID <> "" then
     lcl_statusid = CLng(iStatusID)
  else
     lcl_statusid = 0
  end if

  sSQL = "SELECT statusid, statusname "
  sSQL = sSQL & " FROM egov_mappoints_statuses "
  sSQL = sSQL & " WHERE orgid = " & iOrgID

  if iLimitToActiveOnly then
     sSQL = sSQL & " AND isActive = 1 "
  end if

  set oMPStatuses = Server.CreateObject("ADODB.Recordset")
  oMPStatuses.Open sSQL, Application("DSN"), 3, 1

  if not oMPStatuses.eof then
     do while not oMPStatuses.eof

        if oMPStatuses("statusid") = lcl_statusid then
           lcl_checked_status = " selected=""selected"""
        else
           lcl_checked_status = ""
        end if

        response.write "  <option value=""" & oMPStatuses("statusid") & """" & lcl_checked_status & ">" & oMPStatuses("statusname") & "</option>" & vbcrlf

        oMPStatuses.movenext
     loop
  end if

  oMPStatuses.close
  set oMPStatuses = nothing

end sub

'------------------------------------------------------------------------------
sub getMapPointStatusInfo(ByVal p_mappointid, ByRef lcl_statusid, ByRef lcl_statusname)

  lcl_statusid   = 0
  lcl_statusname = ""

  if p_mappointid <> "" then
     sSQL = "SELECT mp.statusid, mps.statusname "
     sSQL = sSQL & " FROM egov_mappoints mp "
     sSQL = sSQL &      " LEFT OUTER JOIN egov_mappoints_statuses mps ON mp.statusid = mps.statusid "
     sSQL = sSQL & " WHERE mp.mappointid = " & p_mappointid

     set oGetMPStatusInfo = Server.CreateObject("ADODB.Recordset")
     oGetMPStatusInfo.Open sSQL, Application("DSN"), 3, 1

     if not oGetMPStatusInfo.eof then
        lcl_statusid   = oGetMPStatusInfo("statusid")
        lcl_statusname = oGetMPStatusInfo("statusname")
     end if

  end if

  oGetMPStatusInfo.close
  set oGetMPStatusInfo = nothing

end sub

'------------------------------------------------------------------------------
function getMapPointStatusName(iStatusID)

  lcl_return = ""

  if iStatusID <> "" then
     sSQL = "SELECT statusname "
     sSQL = sSQL & " FROM egov_mappoints_statuses "
     sSQL = sSQL & " WHERE statusid = " & iStatusID

     set oGetMPStatusName = Server.CreateObject("ADODB.Recordset")
     oGetMPStatusName.Open sSQL, Application("DSN"), 3, 1

     if not oGetMPStatusName.eof then
        lcl_return = oGetMPStatusName("statusname")
     end if

     oGetMPStatusName.close
     set oGetMPStatusName = nothing

  end if

  getMapPointStatusName = lcl_return

end function

'------------------------------------------------------------------------------
sub updateMapPointValues(iMapPointTypeID)

  if iMapPointTypeID <> "" then
     sSQL = "SELECT mp_fieldid, fieldtype, fieldname, displayInResults, resultsOrder "
     sSQL = sSQL & " FROM egov_mappoints_types_fields "
     sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

     set oGetMPTFields = Server.CreateObject("ADODB.Recordset")
     oGetMPTFields.Open sSQL, Application("DSN"), 3, 1

     if not oGetMPTFields.eof then
        do while not oGetMPTFields.eof

           lcl_fieldtype        = "NULL"
           lcl_fieldname        = "NULL"
           lcl_displayInResults = 0
           lcl_resultsOrder     = 1

           if oGetMPTFields("fieldtype") <> "" then
              lcl_fieldtype = "'" & dbsafe(oGetMPTFields("fieldtype")) & "'"
           end if

           if oGetMPTFields("fieldname") <> "" then
              lcl_fieldname = "'" & dbsafe(oGetMPTFields("fieldname")) & "'"
           end if

           if oGetMPTFields("displayInResults") then
              lcl_displayInResults = 1
           end if

           if oGetMPTFields("resultsOrder") <> "" then
              lcl_resultsOrder = oGetMPTFields("resultsOrder")
           end if


           sSQL = "UPDATE egov_mappoints_values SET "
           sSQL = sSQL & " fieldtype = "        & lcl_fieldtype        & ", "
           sSQL = sSQL & " fieldname = "        & lcl_fieldname        & ", "
           sSQL = sSQL & " displayInResults = " & lcl_displayInResults & ", "
           sSQL = sSQL & " resultsOrder = "     & lcl_resultsOrder
           sSQL = sSQL & " WHERE mp_fieldid = " & oGetMPTFields("mp_fieldid")

           set oUpdateMPValues = Server.CreateObject("ADODB.Recordset")
           oUpdateMPValues.Open sSQL, Application("DSN"), 3, 1

           set oUpdateMPValues = nothing

           oGetMPTFields.movenext
        loop
     end if

     oGetMPTFields.close
     set oGetMPTFields = nothing

  end if

end sub
%>