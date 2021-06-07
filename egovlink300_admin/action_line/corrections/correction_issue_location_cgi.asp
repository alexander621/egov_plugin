<!-- #include file="../../includes/common.asp" //-->
<!-- #include file="../action_line_global_functions.asp" //-->
<!-- #include file="correction_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: CORRECTION_CONTACT_INFO_CGI.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/12/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE CONTACT INFORMATION
'
' MODIFICATION HISTORY
' 1.0	 02/12/07	 John Stullenberger - INITIAL VERSION
' 2.0  10/30/07  David Boyer	- Added "Large Address" and "Validate Address" features
'
'------------------------------------------------------------------------------
' CALL ROUTINE TO SAVE ISSUE LOCATION INFORMATION
Call subSaveIssueInfo()

'------------------------------------------------------------------------------
Sub subSaveIssueInfo()

'GET INFORMATION FROM DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.CursorLocation = 3
	sSQL = "SELECT streetnumber, streetprefix, streetaddress, streetsuffix, streetdirection, streetunit, sortstreetname, city, state, zip, "
 sSQL = sSQL & " comments, latitude, longitude, validstreet, county, parcelidnumber, listedowner, residenttype, legaldescription, "
 sSQL = sSQL & " registereduserid "
 sSQL = sSQL & " FROM egov_action_response_issue_location "
 sSQL = sSQL & " WHERE actionrequestresponseid= " & request("requestid")
 sSQL = sSQL & " AND excludefromactionline = 0 "
	oSave.Open sSQL, Application("DSN"), 1, 2

'UPDATE INFORMATION
	If NOT oSave.EOF Then
  		Dim sNumber, sAddress, sLatitude, sLongitude, sLogEntry, sValidStreet

		 'CHECK TO SEE IF USING ADDRESS FROM LIST OR IF CHOOSING CUSTOM ADDRESS
  		if trim(request.form("ques_issue2")) = "" then
       if OrgHasFeature("large address list") then
          GetAddressInfoLarge request("residentstreetnumber"), request("streetaddress"), sNumber, sPrefix, sAddress, sSuffix, sDirection, _
                              sLatitude, sLongitude, sCity, sState, sZip, sCounty, sParcelID, sListedOwner, sLegalDescription, _
                              sResidentType, sRegisteredUserID, sValidStreet
       else
   		    	GetAddressInfo request("streetaddress"), sNumber, sPrefix, sAddress, sSuffix, sDirection, sLatitude, sLongitude, sCity, sState, _
                         sZip, sCounty, sParcelID, sListedOwner, sLegalDescription, sResidentType, sRegisteredUserID, sValidStreet
       end if
   	else
		     if CLng(request("streetaddress")) <> CLng(0) then
		       'Handle the dropdown addresses - These should have the residentaddressid as the selected value
   		    	GetAddressInfo request("ques_issue2"), sNumber, sPrefix, sAddress, sSuffix, sDirection, sLatitude, sLongitude, sCity, sState, _
                         sZip, sCounty, sParcelID, sListedOwner, sLegalDescription, sResidentType, sRegisteredUserID, sValidStreet
	 	    else
		       'they selected Other address not listed
'          BreakOutAddress request("ques_issue2"), sStreetNumber, sStreetName

'          GetAddressInfoLarge sStreetNumber, sStreetName, sNumber, sPrefix, sAddress, sSuffix, sDirection, sLatitude, sLongitude

          sNumber    = ""
          sPrefix    = ""
          sAddress   = request("ques_issue2")
          sSuffix    = ""
          sDirection = ""
          sLatitude  = 0.00
          sLongitude = 0.00
		     end if
   	end if

 		'COMPARE VALUES - IF CHANGED UPDATE AND LOG
    lcl_original_address     = buildStreetAddress(oSave("streetnumber"), oSave("streetprefix"), oSave("streetaddress"), oSave("streetsuffix"), oSave("streetdirection"))
    lcl_sAddress_street_name = buildStreetAddress(sNumber, sPrefix, sAddress, sSuffix, sDirection)

   'Compare the original address to the new address
   	if trim(lcl_original_address) <> (trim(lcl_sAddress_street_name)) Then
   			'Log the changes
    			If sLogEntry = "" Then
				     	sLogEntry = "Edit Issue Location:<br>"
          sLogEntry = sLogEntry & "Address: """     & trim(lcl_original_address) & """ changed to """ & trim(lcl_sAddress_street_name) & """"
    			Else
     					sLogEntry = sLogEntry & "<br />Address: """ & trim(lcl_original_address) & """ changed to """ & trim(lcl_sAddress_street_name) & """"
    			End If

   			'Save
    			oSave("streetnumber")    = trim(sNumber)
       oSave("streetprefix")    = trim(sPrefix)
    			oSave("streetaddress")   = trim(sAddress)
       oSave("streetsuffix")    = trim(sSuffix)
       oSave("streetdirection") = trim(sDirection)

       if NOT isnull(sLatitude) then
          oSave("latitude") = sLatitude
       else
          oSave("latitude") = 0.00
       end if

       if NOT isnull(sLongitude) then
          oSave("longitude") = sLongitude
       else
          oSave("longitude") = 0.00
       end if
   	End If

	 	'CHECK FOR DIFFERENT VALUES AND SAVE
  		For Each oColumn in oSave.Fields
			    'CHOOSE OPERATION BASED ON COLUMN NAME
     			Select Case oColumn.Name

       			Case "streetnumber","streetaddress","latitude","longitude","validstreet","streetprefix","streetsuffix","streetdirection","sortstreetname","registereduserid","residenttype"
          				'SKIP PROCESSING AS PROCESSED ABOVE
        		Case Else
          				'COMPARE VALUES CITY,STATE,ZIP,COMMENTS - IF CHANGED UPDATE AND LOG
            			If trim(oColumn.Value & " ") <> Trim(request(oColumn.Name)) Then
              				sFieldName = GetFieldDisplayName(oColumn.Name)

             				'LOG CHANGES
              				If sLogEntry = "" Then
               						sLogEntry = "Edit Location Information:<br />"
                     sLogEntry = sLogEntry & sFieldName & ": """ & trim(oColumn.Value) & """ changed to """ & Trim(request(oColumn.Name)) & """"
             					Else
                					sLogEntry = sLogEntry & "<br />"
                     sLogEntry = sLogEntry & sFieldName & ": """ & trim(oColumn.Value) & """ changed to """ & Trim(request(oColumn.Name)) & """"
              				End If
           				End If

          				'SAVE
           				oSave(oColumn.Name) = Trim(request(oColumn.Name))
        End Select
    Next

   'Retrieve the remaining fields
    oSave("validstreet")      = request("validstreet")
    oSave("city")             = request("city")
    oSave("state")            = request("state")
    oSave("zip")              = request("zip")
    oSave("streetunit")       = request("streetunit")
    oSave("county")           = request("county")
    oSave("parcelidnumber")   = request("parcelidnumber")
    oSave("listedowner")      = request("listedowner")
    oSave("residenttype")     = request("residenttype")
    oSave("legaldescription") = request("legaldescription")

    if request("registereduserid") = "" then
       oSave("registereduserid") = 0
    else
       oSave("registereduserid") = request("registereduserid")
    end if

   'Re-build the SortStreetName --------------------------------
    sSortStreetName = trim(oSave("streetaddress"))

    if trim(oSave("streetsuffix")) <> "" then
       if sSortStreetName <> "" then
          sSortStreetName = sSortStreetName & " " & oSave("streetsuffix")
       else
          sSortStreetName = oSave("streetsuffix")
       end if
    end if

    if trim(oSave("streetdirection")) <> "" then
       if sSortStreetName <> "" then
          sSortStreetName = sSortStreetName & " " & oSave("streetdirection")
       else
          sSortStreetName = oSave("streetdirection")
       end if
    end if

    if trim(oSave("streetprefix")) <> "" then
       if sSortStreetName <> "" then
          sSortStreetName = sSortStreetName & " " & oSave("streetprefix")
       else
          sSortStreetName = oSave("streetprefix")
       end if
    end if

    oSave("sortstreetname") = sSortStreetName
   '------------------------------------------------------------

 		'SAVE CHANGES
  		oSave.Update

End If

'CLOSE RECORDSET
oSave.Close
Set oSave = Nothing

'If change were made log them
 if sLogEntry <> "" then
  	'Record in log the save activity
  		AddCommentTaskComment sLogEntry, sExternalMsg, request("status"), request("requestid"), session("userid"), session("orgid"), request("substatus"), "", ""
 end if

 response.redirect "../action_respond.asp?control=" & request("requestid") & "&r=save&status="&request("status")

end sub

'------------------------------------------------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

'------------------------------------------------------------------------------
Function DisplayContactMethod(iValue)

	sSQL = "SELECT * FROM egov_contactmethods WHERE rowid='" & iValue & "'"

	Set oMethods = Server.CreateObject("ADODB.Recordset")
	oMethods.Open sSQL, Application("DSN") , 3, 1
	
	If NOT oMethods.EOF Then
		iReturnValue = oMethods("contactdescription") 
	Else
		iReturnValue = "NOT SPECIFIED"
	End If


	Set oMethods = Nothing
	
	DisplayContactMEthod = iReturnValue
	
End Function

'------------------------------------------------------------------------------
Function GetFieldDisplayName(sValue)

	sReturnValue = sValue

	arrNames = Array("residentstreetnumber","streetaddress","streetunit","city","state","comments","zip","latitude","longitude","county","parcelidnumber","listedowner","residenttype","legaldescription" )
	arrDisplayNames = Array("Street Number","Street Address","Unit","City","State","Comments","Zip","Latitude","Longitude","County","Parcel ID","Listed Owner","Resident Type","Legal Description" )

	' LOOP THRU ARRAYS AND FIND MATCH
	For iLoop = 0 to UBOUND(arrNames)
   		If cstr(trim(arrNames(iLoop))) = cstr(Trim(sValue)) Then
     			sReturnValue = arrDisplayNames(iLoop)
     			Exit For
   		End If
 Next

	GetFieldDisplayName = sReturnValue

End Function

sub dtb_debug(p_value)
  sSQLi = "INSERT INTO my_table_dtb(notes) VALUES('" & replace(p_value,"'","''") & "')"
 	set rsi = Server.CreateObject("ADODB.Recordset")
	 rsi.Open sSQLi, Application("DSN"), 3, 1

end sub
%>