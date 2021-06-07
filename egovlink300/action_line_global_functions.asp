<%
'------------------------------------------------------------------------------
sub DisplayLargeAddressList( p_orgid, sResidenttype )
dim sSql, oAddressList

	sSQL = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & p_orgid
 sSQL = sSQL & " AND excludefromactionline = 0 "
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " ORDER BY sortstreetname "

set oAddressList = Server.CreateObject("ADODB.Recordset")
oAddressList.Open sSQL, Application("DSN"), 3, 1

if NOT oAddressList.EOF then
 		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value="""" size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
 		response.write "<select name=""skip_address"" id=""skip_address"">" & vbcrlf
 		response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf

 		do while NOT oAddressList.eof
      lcl_street_name = buildStreetAddress("", oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

   			response.write "  <option value=""" & lcl_street_name & """>" & lcl_street_name & "</option>" & vbcrlf

   			oAddressList.MoveNext
 		loop

   response.write "</select>&nbsp;" & vbclrf
   response.write "<input type=""button"" class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
end if

oAddressList.Close
set oAddressList = nothing 

end sub

'------------------------------------------------------------------------------
sub DisplayAddress( p_orgid, sResidenttype )
	dim sSql, oAddressList

	sSQL = "SELECT residentaddressid, sortstreetname, isnull(residentstreetnumber,'') as residentstreetnumber, "
 sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & p_orgid
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " ORDER BY sortstreetname, CAST(residentstreetnumber AS INT)"

 set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

	response.write "<select name=""skip_address"" id=""skip_address"">" & vbcrlf
	response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf
		
 do while NOT oAddressList.eof
    lcl_street_name = buildStreetAddress(oAddressList("residentstreetnumber"), oAddressList("residentstreetprefix"), oAddressList("residentstreetname"), oAddressList("streetsuffix"), oAddressList("streetdirection"))

  		response.write "  <option value="""  & oAddressList("residentaddressid")  & """>" & lcl_street_name & "</option>" & vbcrlf
	
  		oAddressList.MoveNext
	loop

	response.write "</select>" & vbcrlf

	oAddressList.Close
	Set oAddressList = Nothing 

End Sub 

'------------------------------------------------------------------------------
Function DisplayAddressNumber( p_orgid, sResidenttype, blninputtype  )
	Dim sSql, oAddressList
	
	sSQL = "SELECT DISTINCT Cast(residentstreetnumber as int) as residentstreetnumber "
 sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid=" & p_orgid
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND residentstreetnumber is not null "
 sSQL = sSQL & " ORDER BY CAST(residentstreetnumber AS INT)"
	set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

	If clng(blninputtype) = 0 Then
		response.write "<select name=""ques_issue1"">" & vbcrlf
	Else
		response.write "<select name=""skip_addressnumber"" onChange=""document.frmRequestAction.ques_issue1.value=document.frmRequestAction.skip_addressnumber[document.frmRequestAction.skip_addressnumber.selectedIndex].value;"">" & vbcrlf
	End If 

	response.write "  <option value="" "">Please select an address number...</option>" & vbcrlf
		
	Do While NOT oAddressList.EOF 
		response.write "  <option value=""" &  oAddressList("residentstreetnumber") & """>" & oAddressList("residentstreetnumber") & "</option>" & vbcrlf
		oAddressList.MoveNext
	Loop

	response.write "</select>" & vbcrlf

	oAddressList.close
	Set oAddressList = Nothing 

End Function

'-- This is the new jQuery "display large address list" -----------------------
sub displayLargeAddressList_new(p_orgid, p_street_number, p_prefix, p_street_name, p_suffix, p_direction)
 dim sSql, oAddressList
 dim lcl_streetnumber, lcl_prefix, lcl_streetname, lcl_suffix, lcl_direction, lcl_compare_address

 lcl_streetnumber    = p_street_number
 lcl_prefix          = p_prefix
 lcl_streetname      = p_street_name
 lcl_suffix          = p_suffix
 lcl_direction       = p_direction
 lcl_compare_address = buildStreetAddress("", lcl_prefix, lcl_streetname, lcl_suffix, lcl_direction)

	sSQL = "SELECT DISTINCT sortstreetname, "
 sSQL = sSQL & " ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, "
 sSQL = sSQL & " ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & p_orgid
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " ORDER BY sortstreetname "

 set oAddressList = Server.CreateObject("ADODB.Recordset")
 oAddressList.Open sSQL, Application("DSN"), 3, 1

 if not oAddressList.eof then
  		'response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" onchange=""save_address();"" /> &nbsp; " & vbcrlf
 	 	'response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkAddressButtons()"">" & vbcrlf
  		response.write "<input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
 	 	response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
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
    'response.write "<input type=""button"" id=""validateAddress"" class=""button"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
    response.write "<input type=""button"" id=""validateAddress"" value=""Validate Address"" onclick=""checkAddress( 'CheckResults', 'no');"" />" & vbcrlf
 else
    'response.write "<table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""font-size:9pt"">" & vbcrlf
    'response.write "  <tr>" & vbcrlf
    'response.write "      <td>Street Number:</td>" & vbcrlf
    'response.write "      <td><input type=""text"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" /></td>" & vbcrlf
 	 	'response.write "  </tr>" & vbcrlf
    'response.write "  <tr>" & vbcrlf
    'response.write "      <td>Street Name:</td>" & vbcrlf
    'response.write "      <td><input type=""text"" name=""streetaddress"" id=""streetaddress"" size=""30"" maxlength=""84"" /></td>" & vbcrlf
    'response.write "  </tr>" & vbcrlf
    'response.write "</table>" & vbcrlf
    response.write "large address list<br />" & vbcrlf
  		response.write "<input type=""hidden"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" /> &nbsp; " & vbcrlf
    response.write "<input type=""hidden"" name=""streetaddress"" id=""streetaddress"" value=""0000"" />" & vbcrlf
 	 	'response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
  		'response.write "  <option value=""0000"">No addresses available</option>" & vbcrlf
    'response.write "</select>" & vbcrlf

 end if

 oAddressList.close
 set oAddressList = nothing

end sub

'-- This is the new display address list --------------------------------------
function displayAddress_new(p_orgid, p_street_number, p_street_name)
	
	dim sNumber, oAddressList, blnFound
 dim lcl_streetnumber, lcl_streetname, lcl_new_street_name

 lcl_streetnumber    = p_street_number
 lcl_streetname      = p_street_name
 lcl_new_street_name = buildStreetAddress(lcl_streetnumber, "", lcl_streetname, "", "")

'Get list of addresses
	sSQL = "SELECT residentaddressid, "
 sSQL = sSQL & " isnull(residentstreetnumber,'') as residentstreetnumber, "
 sSQL = sSQL & " residentstreetprefix, "
 sSQL = sSQL & " residentstreetname, "
 sSQL = sSQL & " streetsuffix, "
 sSQL = sSQL & " streetdirection, "
 sSQL = sSQL & " isnull(latitude,0.00) as latitude, "
 sSQL = sSQL & " isnull(longitude,0.00) as longitude "
 sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid=" & p_orgid
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND residentstreetname is not null "
 sSQL = sSQL & " ORDER BY sortstreetname, residentstreetprefix"

	set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSQL, Application("DSN"), 3, 1

 if not oAddressList.eof then
    response.write "<input type=""hidden"" name=""residentstreetnumber"" id=""residentstreetnumber"" value="""" />" & vbcrlf
   	'response.write "<select name=""streetaddress"" id=""streetaddress"" onchange=""save_address();checkImportAddressBtn();"">" & vbcrlf
   	response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
   	response.write "  <option value=""0000"">Choose street from dropdown</option>" & vbcrlf
	
   	do while not oAddressList.eof
      'Build the original full street address
       lcl_original_street_name = buildStreetAddress(oAddressList("residentstreetnumber"), _
                                                     oAddressList("residentstreetprefix"), _
                                                     oAddressList("residentstreetname"), _
                                                     oAddressList("streetsuffix"), _
                                                     oAddressList("streetdirection")_
                                                    )

       if UCASE(lcl_streetname) = UCASE(oAddressList("residentstreetname")) then
          sSelected = " selected=""selected"""
       else
          sSelected = ""
       end if

      	response.write "  <option value=""" & oAddressList("residentaddressid") & """" & sSelected & ">" & lcl_original_street_name & "</option>" & vbcrlf

  	   	oAddressList.MoveNext
   	loop

   	response.write "</select>" & vbcrlf

    if CityHasGeopointAddresses( p_orgid, "R" ) Then 
       response.write "&nbsp; <input type=""button"" name=""btnMap"" id=""btnMap"" class=""button"" value=""View on Map"" onclick=""ShowMap();"" />" & vbcrlf
    end if
 else
    response.write "small address list<br />" & vbcrlf
    response.write "<input type=""hidden"" name=""residentstreetnumber"" id=""residentstreetnumber"" value=""" & lcl_streetnumber & """ size=""8"" maxlength=""10"" />" & vbcrlf
    response.write "<input type=""hidden"" name=""streetaddress"" id=""streetaddress"" value=""0000"" />" & vbcrlf
   	'response.write "<select name=""streetaddress"" id=""streetaddress"">" & vbcrlf
   	'response.write "  <option value=""0000"">No addresses available</option>" & vbcrlf
    'response.write "</select>" & vbcrlf
 end if

	oAddressList.close
	set oAddressList = nothing

end function

'------------------------------------------------------------------------------
function CityHasGeopointAddresses( iOrgId, sResidenttype )
	dim sSql, oGeoPoints, lcl_return

	lcl_return = false

	sSQL = "SELECT count(latitude) as geopoints "
 sSQL = sSQL & " FROM egov_residentaddresses "
 sSQL = sSQL & " WHERE orgid = " & iOrgId
 sSQL = sSQL & " AND excludefromactionline = 0 "
 sSQL = sSQL & " AND latitude is not NULL "

	set oGeoPoints = Server.CreateObject("ADODB.Recordset")
	oGeoPoints.Open sSQL, Application("DSN"), 0, 1

	if not oGeoPoints.eof then
		 if clng(oGeoPoints("geopoints")) > clng(0) then
   			lcl_return = true
   end if
 end if

	oGeoPoints.close
	set oGeoPoints = nothing 

 CityHasGeopointAddresses = lcl_return

end function

'------------------------------------------------------------------------------
function formatActivityLogComment(iComment)
  lcl_return = ""

  if iComment <> "" then
     lcl_return = iComment
     lcl_return = replace(lcl_return,"default_novalue","")
     lcl_return = replace(lcl_return,chr(10),"<br />")
     lcl_return = replace(lcl_return,chr(13),"")
     lcl_return = replace(lcl_return,"</p><br /><p>","</p><p>")
     lcl_return = replace(lcl_return,"<br /></p><p><br />","</p><p>")
     lcl_return = replace(lcl_return,"<p><br /><b>","<p><b>")
     lcl_return = replace(lcl_return,"</b><br><br />","</b><br />")
  end if

  formatActivityLogComment = lcl_return

end function


'------------------------------------------------------------------------------
function checkOrgForm( ByVal p_orgid, ByVal p_column_name, ByVal p_formid )
  Dim sSql, oCheckOrgForm, lcl_return, lcl_column_name, lcl_formid, lcl_orgformvalue
 'Columns are found on the "org properties" screen and are stored on the "Organizations" table.
 'The fields on the screen:
 ' - Calendar Req Form #: [column name: OrgRequestCalForm]
 ' - Class Evaluation Form: [column name: EvaluationFormID]
 ' - Facility Survey Form: [column name: facilitysurveyformid]
 ' - Rental Survey Form: [column: rentalsurveyformid]
 
  lcl_return       = False
  lcl_orgformvalue = ""

  if trim(p_column_name) <> "" then
     lcl_column_name = p_column_name

     if p_formid <> "" then
        lcl_formid = p_formid

        if isnumeric(lcl_formid) then
           sSQL = "SELECT " & p_column_name & " AS orgformvalue "
           sSQL = sSQL & " FROM organizations "
           sSQL = sSQL & " WHERE orgid = " & p_orgid

          	set oCheckOrgForm = Server.CreateObject("ADODB.Recordset")
          	oCheckOrgForm.Open sSQL, Application("DSN"), 3, 1

           if not oCheckOrgForm.eof then
              lcl_orgformvalue = oCheckOrgForm("orgformvalue")

              if lcl_orgformvalue <> "" then
                 if lcl_orgformvalue = p_formid then
                    lcl_return = True
                 end if
              end if

           end if
        end if

        oCheckOrgForm.close
        set oCheckOrgForm = nothing

     end if
  end if

  checkOrgForm = lcl_return

end function


'------------------------------------------------------------------------------
function getRequestDueDate( ByVal iRequestID)

  lcl_return = ""

  if iRequestID <> "" then
     sSQL = "SELECT due_date "
     sSQL = sSQL & " FROM egov_actionline_requests "
     sSQL = sSQL & " WHERE action_autoid = '" & iRequestID & "'"

     set oGetDueDate = Server.CreateObject("ADODB.Recordset")
     oGetDueDate.Open sSQL, Application("DSN"), 3, 1

     if not oGetDueDate.eof then
        lcl_return = oGetDueDate("due_date")
     end if

     oGetDueDate.close
     set oGetDueDate = nothing

  end if

  getRequestDueDate = lcl_return

end function

'------------------------------------------------------------------------------
function displayMobileOptions(iFormID, iDBColumnName)
   dim lcl_return, sSQL, sFormID, sDBColumnName

   lcl_return    = false
   sFormID       = 0
   sDBColumnName = ""
   sSQL          = ""

   if iFormID <> "" then
      if not containsApostrophe(iFormID) then
         sFormID = clng(iFormID)
      end if
   end if

   if iDBColumnName <> "" then
      sDBColumnName = dbsafe(iDBColumnName)
      sDBColumnName = lcase(sDBColumnName)

      sSQL = "SELECT isnull(" & sDBColumnName & ",0) as mobileoption "
      sSQL = sSQL & " FROM egov_action_request_forms "
      sSQL = sSQL & " WHERE action_form_id = '" & sFormID & "'"

     set oGetMobileOption = Server.CreateObject("ADODB.Recordset")
     oGetMobileOption.Open sSQL, Application("DSN"), 3, 1

     if not oGetMobileOption.eof then
        lcl_return = oGetMobileOption("mobileoption")
     end if

     oGetMobileOption.close
     set oGetMobileOption = nothing
   end if

   displayMobileOptions = lcl_return

end function

'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  dim lcl_return, lcl_orgid, lcl_dmid

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
        lcl_return = "Does not exist..."
     elseif iSuccess = "ERROR" then
        lcl_return = "ERROR"
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "') "
  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 3, 1

end sub
%>
