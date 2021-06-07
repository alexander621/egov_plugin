<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CORRECTION_CONTACT_INFO_CGI.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/12/07
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE CONTACT INFORMATION
'
' MODIFICATION HISTORY
' 1.0	02/12/07	JOHN STULLENBERGER - INITIAL VERSION
'
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------


' CALL ROUTINE TO SAVE CONTACT INFORMATION
Call subSaveContactInfo()


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'------------------------------------------------------------------------------------------------------------
' SUB SUBSAVECONTACTINFO()
'------------------------------------------------------------------------------------------------------------
Sub subSaveContactInfo()

	' GET INFORMATION FROM DATABASE
	Set oSave = Server.CreateObject("ADODB.Recordset")
	oSave.CursorLocation = 3
	sSQL = "SELECT streetnumber,streetaddress,city,state,comments,zip,latitude,longitude FROM egov_action_response_issue_location WHERE actionrequestresponseid= " & request("irequestid")
	oSave.Open sSQL, Application("DSN"), 1, 2

	' UPDATE INFORMATION
	If NOT oSave.EOF Then

		Dim sNumber, sAddress, sLatitude, sLongitude, sLogEntry

		
		' CHECK TO SEE IF USING ADDRESS FROM LIST OR IF CHOOSING CUSTOM ADDRESS
		If trim(request.form("ques_issue2")) = "" Then
			GetAddressInfo request("select_address"), sNumber, sAddress, sLatitude, sLongitude

			' HANDLE LATITUDE AND LONGITUDE
			oSave("latitude") = sLatitude
			oSave("longitude") = sLongitude

		Else
			
			sNumber = ""
			sAddress = dbsafe(trim(request.form("ques_issue2")))

			' HANDLE LATITUDE AND LONGITUDE
			oSave("latitude") = NULL
			oSave("longitude") = NULL
		End If


		' COMPARE VALUES - IF CHANGED UPDATE AND LOG
		If trim(oSave("streetnumber") & " " & oSave("streetaddress")) <> (Trim(sNumber) & " " & Trim(sAddress)) Then

			' LOG CHANGES
			If sLogEntry = "" Then
					sLogEntry = "Edit Issue Location: " & sLogEntry & chr(34) & Trim(oSave("streetnumber") & " " & oSave("streetaddress"))  & chr(34) &  " changed to " & chr(34) &  Trim(Trim(sNumber) & " " &  Trim(sAddress)) & chr(34) 
			Else
					sLogEntry = sLogEntry & "," & chr(34) & Trim(oSave("streetnumber") & " " & oSave("streetaddress")) & chr(34) &  " changed to " & chr(34) &  Trim(Trim(sNumber) & " " & Trim(sAddress)) & chr(34) 
			End If

			' SAVE
			oSave("streetnumber") = Trim(sNumber)
			oSave("streetaddress") = Trim(sAddress)

		End If

		
		' CHECK FOR DIFFERENT VALUES AND SAVE
		For Each oColumn in oSave.Fields

			'CHOOSE OPERATION BASED ON COLUMN NAME
			Select Case oColumn.Name

			Case "streetnumber","streetaddress","latitude","longitude"
				'SKIP PROCESSING AS PROCESSED ABOVE
				
			Case Else

				' COMPARE VALUES CITY,STATE,ZIP,COMMENTS - IF CHANGED UPDATE AND LOG
				If trim(oColumn.Value & " ") <> Trim(request(oColumn.Name)) Then

					' LOG CHANGES
					If sLogEntry = "" Then
						sLogEntry = "Edit Location Information: " & sLogEntry & chr(34) & trim(oColumn.Value)  & chr(34) &  " changed to " & chr(34) &  Trim(request(oColumn.Name)) & chr(34) 
					Else
						sLogEntry = sLogEntry & "," & chr(34) &  trim(oColumn.Value)  & chr(34) &  " changed to " & chr(34) &  Trim(request(oColumn.Name)) & chr(34) 
					End If

				End If
				
				' SAVE
				oSave(oColumn.Name) = Trim(request(oColumn.Name))

			End Select



		Next

		
		' SAVE CHANGES
		oSave.Update
	
	End If

	' CLOSE RECORDSET
	oSave.Close
	Set oSave = Nothing

	response.write sLogEntry

	' RECORD IN LOG THE SAVE ACTIVITY
	Call AddCommentTaskComment(sLogEntry,sExternalMsg,request("status"),request("irequestid"),session("userid"),session("orgid"))


	' RETURN REQUEST PAGE
	response.redirect("correction_issue_location.asp?requestid=" & request("irequestid") & "&r=save&status="&request("status"))

End Sub


'----------------------------------------------------------------------------------------------------------------------
' ADDCOMMENTTASKCOMMENT(SINTERNALMSG,SEXTERNALMSG)
'----------------------------------------------------------------------------------------------------------------------
Function AddCommentTaskComment(sInternalMsg,sExternalMsg,sStatus,iFormID,iUserID,iOrgID)
		
		sSQL = "INSERT egov_action_responses (action_status,action_internalcomment,action_externalcomment,action_userid,action_orgid,action_autoid) VALUES ('" & sStatus & "','" & DBsafe(sInternalMsg) & "','" & DBsafe(sExternalMsg) & "','" & iUserID & "','" & iOrgID & "','" &iFormID & "')"
		Set oComment = Server.CreateObject("ADODB.Recordset")
		oComment.Open sSQL, Application("DSN") , 3, 1
		Set oComment = Nothing

End Function


'----------------------------------------
'  FUNCTION DBSAFE( STRDB )
'----------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function


'------------------------------------------------------------------------------------------------------------
' FUNCTION DISPLAYCONTACTMETHOD(ISELECTED)
'------------------------------------------------------------------------------------------------------------
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


'--------------------------------------------------------------------------------------------------
' Sub GetAddressInfo( sResidentAddressId, ByRef sNumber, ByRef sAddress, ByRef sLatitude, ByRef sLongitude )
'--------------------------------------------------------------------------------------------------
Sub GetAddressInfo( ByVal sResidentAddressId, ByRef sNumber, ByRef sAddress, ByRef sLatitude, ByRef sLongitude )
	Dim sSql, oAddress

	sSql = "Select residentstreetnumber, residentstreetname, residentcity, residentstate, "
	sSql = sSql & " isnull(latitude,0) as latitude, isnull(longitude,0) as longitude From egov_residentaddresses "
	sSql = sSql & " Where residentaddressid = " & sResidentAddressId 

	Set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSQL, Application("DSN"), 0, 1
	
	If Not oAddress.EOF Then 
		sNumber = trim(oAddress("residentstreetnumber"))
		sAddress = oAddress("residentstreetname")
		sCity = oAddress("residentcity")
		sState = oAddress("residentstate")
		sLatitude = oAddress("latitude")
		sLongitude = oAddress("longitude")
	End If 

	oAddress.close
	Set oAddress = Nothing

End Sub 
%>
