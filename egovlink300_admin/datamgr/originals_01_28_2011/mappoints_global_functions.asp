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
sub displayMPTypesFields(iOrgID, iMapPointTypeID, iIsRootAdmin, iIsLimited, iIsDisplayOnly)

  if not iIsDisplayOnly then
     response.write "<div style=""margin-top:20px; margin-bottom:5px;"">" & vbcrlf
     response.write "  <strong>" & lcl_sectiontitle & "</strong><br />" & vbcrlf
     'response.write "  <input type=""button"" name=""reorderButton"" id=""reorderButton"" value=""Maintain Results Field Order"" class=""button"" onclick=""alert('coming soon');"" />" & vbcrlf
     response.write "  <input type=""button"" name=""addMPTField"" id=""addMPTField"" value=""Add Field"" class=""button"" onclick=""addFieldRow();"" />" & vbcrlf
     response.write "</div>" & vbcrlf
  end if

  response.write "<table id=""addFieldTBL"" border=""0"" cellspacing=""0"" cellpadding=""3"" class=""tableadmin"">" & vbcrlf
  response.write "  <tr id=""addFieldRow0"">" & vbcrlf
  response.write "      <th align=""left"" colspan=""2"">Field Name</th>" & vbcrlf
  response.write "      <th>Display In<br />Results</th>" & vbcrlf
  response.write "      <th>Display On<br />Info Page</th>" & vbcrlf
  response.write "      <th>Display<br />Order</th>" & vbcrlf
  response.write "      <th>In Public<br />Search</th>" & vbcrlf
  response.write "      <th>Include<br />""Add a Link""</th>" & vbcrlf
  response.write "      <th>Display as<br />Multi-Line</th>" & vbcrlf
  response.write "      <th>Remove</th>" & vbcrlf
  response.write "  </tr>" & vbcrlf

  iRowCount   = 0
  lcl_bgcolor = "#ffffff"

  sSQL = "SELECT mp_fieldid, "
  sSQL = sSQL & " mappoint_typeid, "
  sSQL = sSQL & " orgid, "
  sSQL = sSQL & " fieldname, "
  sSQL = sSQL & " isnull(fieldtype,'') as fieldtype, "
  sSQL = sSQL & " hasAddLinkButton, "
  sSQL = sSQL & " isMultiLine, "
  sSQL = sSQL & " displayInResults, "
  sSQL = sSQL & " displayInInfoPage, "
  sSQL = sSQL & " resultsOrder, "
  sSQL = sSQL & " inPublicSearch "
  sSQL = sSQL & " FROM egov_mappoints_types_fields "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID
  sSQL = sSQL & " ORDER BY resultsOrder, mp_fieldid "

  set oMPTFields = Server.CreateObject("ADODB.Recordset")
  oMPTFields.Open sSQL, Application("DSN"), 3, 1

  if not oMPTFields.eof then
     do while not oMPTFields.eof

        iRowCount                     = iRowCount + 1
        lcl_bgcolor                   = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")
        lcl_checked_hasAddLinkButton  = isCheckboxChecked(oMPTFields("hasAddLinkButton"))
        lcl_checked_isMultiLine       = isCheckboxChecked(oMPTFields("isMultiLine"))
        lcl_checked_displayInResults  = isCheckboxChecked(oMPTFields("displayInResults"))
        lcl_checked_displayInInfoPage = isCheckboxChecked(oMPTFields("displayInInfoPage"))
        lcl_checked_inPublicSearch    = isCheckboxChecked(oMPTFields("inPublicSearch"))
        lcl_resultsOrder              = oMPTFields("resultsOrder")

        response.write "  <tr id=""addFieldRow" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ align=""center"">" & vbcrlf
        response.write "      <td align=""left"">" & iRowCount & ".</td>" & vbcrlf
        response.write "      <td align=""right"">" & vbcrlf
        response.write "          <input type=""text"" name=""fieldname" & iRowCount & """ id=""fieldname" & iRowCount & """ value=""" & oMPTFields("fieldname") & """ size=""50"" maxlength=""100"" onchange=""clearMsg('fieldname" & iRowCount & "');"" />" & vbcrlf

        if iIsRootAdmin and not iIsLimited then
           response.write "<br /><strong>Field Type: </strong>(code use ONLY)&nbsp;" & vbcrlf
           response.write "<input type=""text"" name=""fieldtype" & iRowCount & """ id=""fieldtype" & iRowCount & """ value=""" & oMPTFields("fieldtype") & """ size=""15"" maxlength=""100"" />" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""displayInResults" & iRowCount & """ id=""displayInResults" & iRowCount & """ value=""1""" & lcl_checked_displayInResults & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""displayInInfoPage" & iRowCount & """ id=""displayInInfoPage" & iRowCount & """ value=""1""" & lcl_checked_displayInInfoPage & " />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""text"" name=""resultsOrder" & iRowCount & """ id=""resultsOrder" & iRowCount & """ value=""" & lcl_resultsOrder & """ size=""3"" maxlength=""5"" />" & vbcrlf
        response.write "      </td>" & vbcrlf
        response.write "      <td>" & vbcrlf
        response.write "          <input type=""checkbox"" name=""inPublicSearch" & iRowCount & """ id=""inPublicSearch" & iRowCount & """ value=""1""" & lcl_checked_inPublicSearch & " />" & vbcrlf
        response.write "      </td>" & vbcrlf

        if oMPTFields("fieldtype") = "" then
           response.write "      <td><input type=""checkbox"" name=""hasAddLinkButton" & iRowCount & """ id=""hasAddLinkButton" & iRowCount & """ value=""1""" & lcl_checked_hasAddLinkButton & " /></td>" & vbcrlf
           response.write "      <td><input type=""checkbox"" name=""isMultiLine" & iRowCount & """ id=""isMultiLine" & iRowCount & """ value=""1""" & lcl_checked_isMultiLine & " /></td>" & vbcrlf
        else
           response.write "      <td><input type=""hidden"" name=""hasAddLinkButton" & iRowCount & """ id=""hasAddLinkButton" & iRowCount & """ value=""0"" /></td>" & vbcrlf
           response.write "      <td><input type=""hidden"" name=""isMultiLine" & iRowCount & """ id=""isMultiLine" & iRowCount & """ value=""0"" /></td>" & vbcrlf
        end if

        response.write "      <td>" & vbcrlf
        response.write "          <input type=""hidden"" name=""mp_fieldid" & iRowCount & """ id=""mp_fieldid" & iRowCount & """ value=""" & oMPTFields("mp_fieldid") & """ />" & vbcrlf

        if not iIsRootAdmin or (iIsRootAdmin AND iIsLimited) then
           response.write "          <input type=""hidden"" name=""fieldtype" & iRowCount & """ id=""fieldtype" & iRowCount & """ value=""" & oMPTFields("fieldtype") & """ size=""20"" maxlength=""100"" />" & vbcrlf
        end if

        if oMPTFields("fieldtype") = "" then
           response.write "          <input type=""checkbox"" name=""deleteField" & iRowCount & """ id=""deleteField" & iRowCount & """ value=""Y"" />" & vbcrlf
           'response.write "          <input type=""button"" name=""deleteButton" & iRowCount & """ id=""deleteButton" & iRowCount & """ value=""Delete"" class=""button"" onclick=""deleteField('" & iRowCount & "');"" />" & vbcrlf
        else
           response.write "          <input type=""hidden"" name=""deleteField" & iRowCount & """ id=""deleteField" & iRowCount & """ value=""N"" />" & vbcrlf
        end if

        response.write "      </td>" & vbcrlf
        response.write "  </tr>" & vbcrlf

        oMPTFields.movenext
     loop
  end if

  response.write "</table>" & vbcrlf
  response.write "<input type=""hidden"" name=""totalFields"" id=""totalFields"" value=""" & iRowCount & """ size=""3"" maxlength=""100"" />" & vbcrlf

  oMPTFields.close
  set oMPTFields = nothing

end sub

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
sub maintainMapPointTypeField(iMPFieldID, iMapPointTypeID, iOrgID, iFieldName, iFieldType, iHasAddLinkButton, iIsMultiLine, _
                              iDisplayInResults, iDisplayInInfoPage, iResultsOrder, iInPublicSearch)

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

  if iHasAddLinkButton <> "" then
     lcl_hasAddLinkButton = iHasAddLinkButton
  else
     lcl_hasAddLinkButton = 0
  end if

  if iIsMultiLine <> "" then
     lcl_isMultiLine = iIsMultiLine
  else
     lcl_isMultiLine = 0
  end if

  if iDisplayInResults <> "" then
     lcl_displayInResults = iDisplayInResults
  else
     lcl_displayInResults = 0
  end if

  if iDisplayInInfoPage <> "" then
     lcl_displayInInfoPage = iDisplayInInfoPage
  else
     lcl_displayInInfoPage = 0
  end if

  if iResultsOrder <> "" then
     lcl_resultsOrder = iResultsOrder
  else
     lcl_resultsOrder = 0
  end if

  if iInPublicSearch <> "" then
     lcl_inPublicSearch = iInPublicSearch
  else
     lcl_inPublicSearch = 0
  end if

 'Determine if the Map-Point is to be added or updated
  if iMPFieldID <> "" then
     sSQL = "UPDATE egov_mappoints_types_fields SET "
     sSQL = sSQL & " fieldname = "         & lcl_fieldname         & ", "
     sSQL = sSQL & " fieldtype = "         & lcl_fieldtype         & ", "
     sSQL = sSQL & " hasAddLinkButton = "  & lcl_hasAddLinkButton  & ", "
     sSQL = sSQL & " isMultiLine = "       & lcl_isMultiLine       & ", "
     sSQL = sSQL & " displayInResults = "  & lcl_displayInResults  & ", "
     sSQL = sSQL & " displayInInfoPage = " & lcl_displayInInfoPage & ", "
     sSQL = sSQL & " resultsOrder = "      & lcl_resultsOrder      & ", "
     sSQL = sSQL & " inPublicSearch = "    & lcl_inPublicSearch
     sSQL = sSQL & " WHERE mp_fieldid = " & iMPFieldID

  else
     sSQL = "INSERT INTO egov_mappoints_types_fields ("
     sSQL = sSQL & "mappoint_typeid,"
     sSQL = sSQL & "orgid,"
     sSQL = sSQL & "fieldname,"
     sSQL = sSQL & "fieldtype,"
     sSQL = sSQL & "hasAddLinkButton,"
     sSQL = sSQL & "isMultiLine,"
     sSQL = sSQL & "displayInResults,"
     sSQL = sSQL & "displayInInfoPage,"
     sSQL = sSQL & "resultsOrder, "
     sSQL = sSQL & "inPublicSearch "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & iMapPointTypeID       & ", "
     sSQL = sSQL & iOrgID                & ", "
     sSQL = sSQL & lcl_fieldname         & ", "
     sSQL = sSQL & lcl_fieldtype         & ", "
     sSQL = sSQL & lcl_hasAddLinkButton  & ", "
     sSQL = sSQL & lcl_isMultiLine       & ", "
     sSQL = sSQL & lcl_displayInResults  & ", "
     sSQL = sSQL & lcl_displayInInfoPage & ", "
     sSQL = sSQL & lcl_resultsOrder      & ", "
     sSQL = sSQL & lcl_inPublicSearch
     sSQL = sSQL & ")"
  end if

	 set oMaintainMapPointTypeField = Server.CreateObject("ADODB.Recordset")
 	oMaintainMapPointTypeField.Open sSQL, Application("DSN"), 3, 1

  set oMaintainMapPointTypeField = nothing

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

 lcl_streetnumber    = p_street_number
 lcl_streetname      = p_street_name
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
          if UCASE(lcl_streetname) = UCASE(oAddressList("residentstreetname")) then
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
sub DisplayLargeAddressList(p_orgid, p_street_number, p_prefix, p_street_name, p_suffix, p_direction)
 dim sSql, oAddressList

 lcl_streetnumber    = p_street_number
 lcl_prefix          = p_prefix
 lcl_streetname      = p_street_name
 lcl_suffix          = p_suffix
 lcl_direction       = p_direction
 lcl_compare_address = buildStreetAddress("", lcl_prefix, lcl_streetname, lcl_suffix, lcl_direction)

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
  		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" onchange=""save_address();"" /> &nbsp; " & vbcrlf
 	 	response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkAddressButtons()"">" & vbcrlf
  		response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf

    do while not oAddressList.eof

      'Build the full street address
       lcl_streetaddress = buildStreetAddress("", oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

      'Determine if the option is selected
     		if UCASE(lcl_streetaddress) = UCASE(lcl_compare_address) then
       			lcl_selected_address = " selected=""selected"""
       else
          lcl_selected_address = ""
     		end if

       response.write "<option value=""" & lcl_streetaddress & """" & lcl_selected_address & ">" & lcl_streetaddress & "</option>" & vbcrlf

    			oAddressList.MoveNext
    loop

    response.write "</select>&nbsp;" & vbcrlf
    response.write "<input type=""button"" id=""validateAddress"" class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
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
sub displayMapPointTypes(iOrgID, iMapPointTypeID, iFeature)

  if iMapPointTypeID <> "" then
     lcl_mappointtypeid = CLng(iMapPointTypeID)
  else
     lcl_mappointtypeid = 0
  end if

  sSQL = "SELECT mappoint_typeid, description "
  sSQL = sSQL & " FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND isActive = 1 "

  if iFeature <> "" AND iFeature <> "mappoints_maint" then
     sSQL = sSQL & " AND UPPER(feature_maintain) = '" & UCASE(iFeature) & "' "
  end if

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
sub maintainMapPointValues(iOrgID, iMapPointTypeID, iMapPointID, iMPFieldID, iMPValueID, iFieldType, iFieldValue)

  lcl_orgid           = iOrgID
  lcl_mappoint_typeid = iMapPointTypeID
  lcl_mappointid      = iMapPointID
  lcl_mp_fieldid      = iMPFieldID
  lcl_mp_valueid      = 0

  if iMPValueID <> "" then
     lcl_mp_valueid = iMPValueID
  end if

  lcl_fieldtype  = formatFieldforInsertUpdate(iFieldType)
  lcl_fieldvalue = formatFieldforInsertUpdate(iFieldValue)

 'If a mp_valueid exists then update the Map-Point Value.  Otherwise, insert it.
  if lcl_mp_valueid > 0 then
     sSQL = "UPDATE egov_mappoints_values SET "
     sSQL = sSQL & " mappoint_typeid = "  & lcl_mappoint_typeid  & ", "
     sSQL = sSQL & " fieldtype = "        & lcl_fieldtype        & ", "
     sSQL = sSQL & " fieldvalue = "       & lcl_fieldvalue
     sSQL = sSQL & " WHERE mp_valueid = " & lcl_mp_valueid

  else

     sSQL = "INSERT INTO egov_mappoints_values ("
     sSQL = sSQL & "orgid, "
     sSQL = sSQL & "mappoint_typeid, "
     sSQL = sSQL & "mappointid, "
     sSQL = sSQL & "mp_fieldid, "
     sSQL = sSQL & "fieldtype, "
     sSQL = sSQL & "fieldvalue "
     sSQL = sSQL & ") VALUES ("
     sSQL = sSQL & lcl_orgid            & ", "
     sSQL = sSQL & lcl_mappoint_typeid  & ", "
     sSQL = sSQL & lcl_mappointid       & ", "
     sSQL = sSQL & lcl_mp_fieldid       & ", "
     sSQL = sSQL & lcl_fieldtype        & ", "
     sSQL = sSQL & lcl_fieldvalue
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
'sub displayMPTypeStatuses(iOrgID, iStatusID, iLimitToActiveOnly)

'  if iStatusID <> "" then
'     lcl_statusid = CLng(iStatusID)
'  else
'     lcl_statusid = 0
'  end if

'  sSQL = "SELECT statusid, statusname "
'  sSQL = sSQL & " FROM egov_mappoints_statuses "
'  sSQL = sSQL & " WHERE orgid = " & iOrgID

'  if iLimitToActiveOnly then
'     sSQL = sSQL & " AND isActive = 1 "
'  end if

'  set oMPStatuses = Server.CreateObject("ADODB.Recordset")
'  oMPStatuses.Open sSQL, Application("DSN"), 3, 1

'  if not oMPStatuses.eof then
'     do while not oMPStatuses.eof

'        if oMPStatuses("statusid") = lcl_statusid then
'           lcl_checked_status = " selected=""selected"""
'        else
'           lcl_checked_status = ""
'        end if

'        response.write "  <option value=""" & oMPStatuses("statusid") & """" & lcl_checked_status & ">" & oMPStatuses("statusname") & "</option>" & vbcrlf

'        oMPStatuses.movenext
'     loop
'  end if

'  oMPStatuses.close
'  set oMPStatuses = nothing

'end sub

'------------------------------------------------------------------------------
'sub getMapPointStatusInfo(ByVal p_mappointid, ByRef lcl_statusid, ByRef lcl_statusname)

'  lcl_statusid   = 0
'  lcl_statusname = ""

'  if p_mappointid <> "" then
'     sSQL = "SELECT mp.statusid, mps.statusname "
'     sSQL = sSQL & " FROM egov_mappoints mp "
'     sSQL = sSQL &      " LEFT OUTER JOIN egov_mappoints_statuses mps ON mp.statusid = mps.statusid "
'     sSQL = sSQL & " WHERE mp.mappointid = " & p_mappointid

'     set oGetMPStatusInfo = Server.CreateObject("ADODB.Recordset")
'     oGetMPStatusInfo.Open sSQL, Application("DSN"), 3, 1

'     if not oGetMPStatusInfo.eof then
'        lcl_statusid   = oGetMPStatusInfo("statusid")
'        lcl_statusname = oGetMPStatusInfo("statusname")
'     end if

'  end if

'  oGetMPStatusInfo.close
'  set oGetMPStatusInfo = nothing

'end sub

'------------------------------------------------------------------------------
'function getMapPointStatusName(iStatusID)

'  lcl_return = ""

'  if iStatusID <> "" then
'     sSQL = "SELECT statusname "
'     sSQL = sSQL & " FROM egov_mappoints_statuses "
'     sSQL = sSQL & " WHERE statusid = " & iStatusID

'     set oGetMPStatusName = Server.CreateObject("ADODB.Recordset")
'     oGetMPStatusName.Open sSQL, Application("DSN"), 3, 1

'     if not oGetMPStatusName.eof then
'        lcl_return = oGetMPStatusName("statusname")
'     end if

'     oGetMPStatusName.close
'     set oGetMPStatusName = nothing

'  end if

'  getMapPointStatusName = lcl_return

'end function

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

'------------------------------------------------------------------------------
sub displayMapPointColors(iMapPointColor)

 'Determine which color is selected
  if iMapPointColor = "blue" then
     lcl_selected_mpcolor_blue   = " selected=""selected"""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = ""
  elseif iMapPointColor = "orange" then
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = " selected=""selected"""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = ""
  elseif iMapPointColor = "pink" then
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = " selected=""selected"""
     lcl_selected_mpcolor_red    = ""
  elseif iMapPointColor = "red" then
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = ""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = " selected=""selected"""
  else
     lcl_selected_mpcolor_blue   = ""
     lcl_selected_mpcolor_green  = " selected=""selected"""
     lcl_selected_mpcolor_orange = ""
     lcl_selected_mpcolor_pink   = ""
     lcl_selected_mpcolor_red    = ""
  end if

  response.write "  <option value=""blue"""   & lcl_selected_mpcolor_blue   & ">Blue</option>"   & vbcrlf
  response.write "  <option value=""green"""  & lcl_selected_mpcolor_green  & ">Green</option>"  & vbcrlf
  response.write "  <option value=""orange""" & lcl_selected_mpcolor_orange & ">Orange</option>" & vbcrlf
  response.write "  <option value=""pink"""   & lcl_selected_mpcolor_pink   & ">Pink</option>"   & vbcrlf
  response.write "  <option value=""red"""    & lcl_selected_mpcolor_red    & ">Red</option>"    & vbcrlf

end sub

'------------------------------------------------------------------------------
function getMapPointTypePointColor(iMapPointTypeID)

  lcl_return = "green"

  sSQL = "SELECT mappointcolor "
  sSQL = sSQL & " FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE mappoint_typeid = " & iMapPointTypeID

  set oGetMPTypeColor = Server.CreateObject("ADODB.Recordset")
  oGetMPTypeColor.Open sSQL, Application("DSN"), 3, 1

  if not oGetMPTypeColor.eof then
     lcl_return = oGetMPTypeColor("mappointcolor")
  end if

  oGetMPTypeColor.close
  set oGetMPTypeColor = nothing

  getMapPointTypePointColor = lcl_return

end function

'------------------------------------------------------------------------------
function getMapPointTypeByFeature(p_orgid, p_featuresearch, p_feature)

  lcl_return        = 0
  lcl_featuresearch = "feature_maintain"

  if p_featuresearch <> "" then
     lcl_featuresearch = p_featuresearch
  end if

  if p_feature <> "" then
     sSQL = "SELECT mappoint_typeid "
     sSQL = sSQL & " FROM egov_mappoints_types "
     sSQL = sSQL & " WHERE UPPER(" & lcl_featuresearch & ") = '" & UCASE(p_feature) & "' "
     sSQL = sSQL & " AND orgid = " & p_orgid

     set oGetMPTID = Server.CreateObject("ADODB.Recordset")
     oGetMPTID.Open sSQL, Application("DSN"), 3, 1

     if not oGetMPTID.eof then
        lcl_return = oGetMPTID("mappoint_typeid")
     end if

     oGetMPTID.close
     set oGetMPTID = nothing

  end if

  getMapPointTypeByFeature = lcl_return

end function

'------------------------------------------------------------------------------
sub GetCityPoint(ByVal p_orgid, ByRef sLat, ByRef sLng )
    sLat = ""
    sLng = ""

   'Get the point to center the map
    sSQL = "SELECT latitude, longitude "
    sSQL = sSQL & " FROM organizations "
    sSQL = sSQL & " WHERE orgid = " & p_orgid

    set oCityPoint = Server.CreateObject("ADODB.Recordset")
    oCityPoint.Open sSQL, Application("DSN"), 3, 1

    if not oCityPoint.eof then
       sLat = oCityPoint("latitude")
       sLng = oCityPoint("longitude")
    end if

    oCityPoint.close
    set oCityPoint = nothing

end sub

'------------------------------------------------------------------------------
sub displayTemplateOptions()

  response.write "  <option value=""""></option>" & vbcrlf

  sSQL = "SELECT mappoint_typeid, "
  sSQL = sSQL & " description "
  sSQL = sSQL & " FROM egov_mappoints_types "
  sSQL = sSQL & " WHERE isTemplate = 1 "
  sSQL = sSQL & " AND isActive = 1 "

  set oGetMPTemplates = Server.CreateObject("ADODB.Recordset")
  oGetMPTemplates.Open sSQL, Application("DSN"), 3, 1

  if not oGetMPTemplates.eof then
     do while not oGetMPTemplates.eof

        response.write "  <option value=""" & oGetMPTemplates("mappoint_typeid") & """>" & oGetMPTemplates("description") & "</option>" & vbcrlf

        oGetMPTemplates.movenext
     loop
  end if

  oGetMPTemplates.close
  set oGetMPTemplates = nothing

end sub

'------------------------------------------------------------------------------
function isCheckboxChecked(iValue)

  lcl_return = ""

  if iValue then
     lcl_return = " checked=""checked"""
  end if

  isCheckboxChecked = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(iValue)

  lcl_return = ""

  if iValue <> "" then
     lcl_return = replace(iValue,"'","''")
  end if

  dbsafe = lcl_return


end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

end sub
%>