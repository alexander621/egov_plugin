<%
Response.AddHeader "Access-Control-Allow-Origin", "*"
%>
<!-- #include file="../includes/common.asp" //-->
<%

set objUpload = Server.CreateObject("Dundas.Upload.2")
objUpload.MaxFileSize = (31457280) ' MAX SIZE OF UPLOAD SPECIFIED IN BYTES, APPX. 30MB
objUpload.SaveToMemory

'for each item in objUpload.Form
	'name = item & ""
	'value = ""
	'for x = 0 to 10
		'on error resume next
		'value = value & objUpload.Form(name)(x) & ","
		'on error goto 0
	'next
	'if right(value,1) = "," then value = left(value, len(value)-1)
	''response.write name & " = " & value & "<br />" & vbcrlf
'next

intOrgID = Track_DBSafe(objUpload.Form("orgid"))
intPermitTypeID = Track_DBSafe(objUpload.Form("permittypeid"))
intUserID = Track_DBSafe(objUpload.Form("userid"))

'INSERT THIS INTO DATABASE TO GET SUBMISSION ID
sSQL = "INSERT INTO egov_permitapplication_submitted (orgid,permittypeid, userid) VALUES(" & intOrgID & "," & intPermitTypeID & "," & intUserID & ")"
'response.write sSQL & vbcrlf
'response.flush
intPermitApplicationID = RunIdentityInsertStatement(sSQL)

'response.write intOrgID



'Need to loop through the possible fields
for x = 1 to 200
	'on error resume next
	strResposne = ""

	intQuestionID = getField(x, "_questionid")
	
	strFieldType = getField(x, "_fieldtype")
	if strFieldType = "address" then
		strStreetNumber = getField(x, "_StreetNumber")
		strStreetName = getField(x, "_StreetName")
		strStreetSuffix = getField(x, "_StreetSuffix")
		strCity = getField(x, "_City")
		strState = getField(x, "_State")
		strZip = getField(x, "_Zip")
		strCounty = getField(x, "_County")

		if strStreetNumber <> "" or strStreetName <> "" or strStreetSuffix <> "" or strCity <> "" or strState <> "" or strZip <> "" or strCounty <> "" then
			'INSERT INTO ADDRESS
			sSQL = "INSERT INTO egov_permitapplication_address (residentstreetnumber,residentstreetname,streetsuffix,residentcity,residentstate,residentzip,county) " _
				& " VALUES('" & strStreetNumber & "','" & strStreetName & "'," _
				& "'" & strStreetSuffix & "','" & strCity & "','" & strState & "','" & strZip & "','" & strCounty & "')"
			'response.write sSQL & vbcrlf
			'response.flush
			strResponse = RunIdentityInsertStatement(sSQL)
		end if
	
	elseif strFieldType = "contact" then
		strFirstName = getField(x, "_FirstName")
		strLastName = getField(x, "_LastName")
		strCompany = getField(x, "_Company")
		strAddress = getField(x, "_Address")
		strCity = getField(x, "_City")
		strState = getField(x, "_State")
		strZip = getField(x, "_Zip")
		strEmail = getField(x, "_Email")
		strPhone = getField(x, "_Phone")
		strCell = getField(x, "_Cell")
		strFax = getField(x, "_Fax")

		if strFirstName <> "" or strLastName <> "" or strCompany <> "" or strAddress <> "" or strCity <> "" _
			or strState <> "" or strZip <> "" or strEmail <> "" or strPhone <> "" or strCell <> "" or strFax <> "" then

			'INSERT INTO CONTACT
			sSQL = "INSERT INTO egov_permitapplication_contacts (address," _
				& "city,state,zip,company,firstname,lastname,email,phone,cell,fax) " _
				& " VALUES('" & strAddress & "'," _
				& "'" & strCity & "','" & strState & "','" & strZip & "','" & strCompany & "'," _
				& "'" & strFirstName & "','" & strLastName & "','" & strEmail & "','" & strPhone & "','" & strCell & "','" & strFax & "')"
			'response.write sSQL & vbcrlf
			'response.flush
			strResponse = RunIdentityInsertStatement(sSQL)
		end if
	else
		'INSET INTO ANSWER
		strResponse = getField(x,"")
	end if

	if strResponse <> "" then
		sSQL = "INSERT INTO egov_permitapplication_answers (permitapplication_submittedid, permitapplicationquestionid,answer) VALUES(" & intPermitApplicationID & "," & intQuestionID & ",'" & strResponse & "')"
		'response.write sSQL & vbcrlf
		'response.flush
		RunSQLStatement(sSQL)
	end if


	'on error goto 0
next

response.write intPermitApplicationID

function getField(intFieldID, strName)

	getField = Track_DBsafe(objUpload.form("customfield" & intFieldID & strName))


end function

'------------------------------------------------------------------------------------------------------------
'  integer SetOrganizationParameters()
'------------------------------------------------------------------------------------------------------------
Function SetOrganizationParameters()
	Dim sSql, oRs, iReturnValue, sProtocol, sCurrent, sServer

	' SET DEFAULT RETURN VALUE
	iReturnValue = CLng(0)

	' BUILD CURRENT URL
	If request.servervariables("HTTPS") = "on" Then
		sProtocol = "https://"
	Else
		sProtocol = "http://"
	End If
	
	sServer = request.servervariables("SERVER_NAME")
	
	' Translate secure payment URL to regular URL for Org lookup
	If LCase(sServer) = "secure.egovlink.com" or LCase(sServer) = "www.egovlink.com" Then
		sCurrent = "http://www.egovlink.com/" & GetVirtualDirectyName()
	Else
		sCurrent = sProtocol & sServer & "/" & GetVirtualDirectyName()
	End if 
	
	' LOOKUP CURRENT URL IN DATABASE
	'sSql = "SELECT * FROM Organizations INNER JOIN TimeZones ON Organizations.OrgTimeZoneID = TimeZones.TimeZoneID WHERE OrgEgovWebsiteURL = '" & sCurrent & "'"
	sSql = "SELECT OrgID, "
	sSql = sSql & " OrgName, "
	sSql = sSql & " OrgPublicWebsiteURL, "
	sSql = sSql & " OrgEgovWebsiteURL, "
	sSql = sSql & " OrgTopGraphicLeftURL, "
	sSql = sSql & " OrgTopGraphicRightURL,"
	sSql = sSql & " OrgWelcomeMessage, "
	sSql = sSql & " OrgActionLineDescription, "
	sSql = sSql & " OrgPaymentDescription, "
	sSql = sSql & " OrgHeaderSize, "
	sSql = sSql & " OrgTagline, "
	sSql = sSql & " OrgPaymentGateway, "
	sSql = sSql & " OrgActionOn, "
	sSql = sSql & " OrgPaymentOn, "
	sSql = sSql & " OrgDocumentOn, "
	sSql = sSql & " OrgCalendarOn, "
	sSql = sSql & " OrgFaqOn, "
	sSql = sSql & " orgVirtualSiteName, "
	sSql = sSql & " OrgActionName, "
	sSql = sSql & " OrgPaymentName, "
	sSql = sSql & " OrgCalendarName, "
	sSql = sSql & " OrgDocumentName, "
	sSql = sSql & " OrgRegistration, "
	sSql = sSql & " OrgRequestCalOn, "
	sSql = sSql & " OrgRequestCalForm, "
	sSql = sSql & " OrgPublicWebsiteTag, "
	sSql = sSql & " OrgEgovWebsiteTag, "
	sSql = sSql & " OrgCustomButtonsOn, "
	sSql = sSql & " gmtoffset, "
	sSql = sSql & " orgDisplayMenu, "
	sSql = sSql & " orgDisplayFooter, "
	sSql = sSql & " orgCustomMenu, "
	sSql = sSql & " defaultphone, "
	sSql = sSql & " defaultemail, "
	sSql = sSql & " defaultstate, "
	sSql = sSql & " defaultcity, "
	sSql = sSql & " defaultzip, "
	sSql = sSql & " separate_index_catalog, "
	sSql = sSql & " orgwaivertext, "
	sSql = sSql & " OrgGoogleAnalyticAccnt "
	sSql = sSql & " FROM organizations O, timezones T "
	sSql = sSql & " WHERE O.OrgTimeZoneID = T.TimeZoneID "
	sSql = sSql & " AND O.OrgEgovWebsiteURL = '" & sCurrent & "'"

	session("OrgParametersSql") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	session("OrgParametersSql") = ""
	
	If Not oRs.EOF Then
		iOrgID             = oRs("OrgID")
		sOrgName           = oRs("OrgName")
		sHomeWebsiteURL    = oRs("OrgPublicWebsiteURL")
		sEgovWebsiteURL    = oRs("OrgEgovWebsiteURL")
		sTopGraphicLeftURL = oRs("OrgTopGraphicLeftURL")
		sTopGraphicRighURL = oRs("OrgTopGraphicRightURL")

		If request.servervariables("HTTPS") = "on" Then
			 'ADJUST FOR PAYMENT URL
			  sTopGraphicLeftURL = Replace(oRs("OrgTopGraphicLeftURL"),"http:","https:")
			  sTopGraphicRighURL = Replace(oRs("OrgTopGraphicRightURL"),"http:","https:")
			sEgovWebsiteURL    = Replace(oRs("OrgEgovWebsiteURL"),"http:","https:")
		End If

		sWelcomeMessage      = oRs("OrgWelcomeMessage")
		sActionDescription   = oRs("OrgActionLineDescription")
		sPaymentDescription  = oRs("OrgPaymentDescription")
		iHeaderSize          = oRs("OrgHeaderSize")
		sTagline             = oRs("OrgTagline")
		iPaymentGatewayID    = oRs("OrgPaymentGateway")
		blnOrgAction         = oRs("OrgActionOn")
		blnOrgPayment        = oRs("OrgPaymentOn")
		blnOrgDocument       = oRs("OrgDocumentOn")
		blnOrgCalendar       = oRs("OrgCalendarOn")
		blnOrgFaq            = oRs("OrgFaqOn")
		sorgVirtualSiteName  = oRs("orgVirtualSiteName")
		sOrgActionName       = oRs("OrgActionName")
		sOrgPaymentName      = oRs("OrgPaymentName")
		sOrgCalendarName     = oRs("OrgCalendarName")
		sOrgDocumentName     = oRs("OrgDocumentName")
		'sOrgFaqName          = oRs("OrgFaqName")
		sOrgRegistration     = oRs("OrgRegistration")
		blnCalRequest        = oRs("OrgRequestCalOn")
		iCalForm             = oRs("OrgRequestCalForm")
		sHomeWebsiteTag      = oRs("OrgPublicWebsiteTag")
		sEgovWebsiteTag      = oRs("OrgEgovWebsiteTag")
		bCustomButtonsOn     = oRs("OrgCustomButtonsOn")
		iTimeOffset          = oRs("gmtoffset")
		blnMenuOn            = oRs("orgDisplayMenu")
		blnFooterOn          =	oRs("orgDisplayFooter")
		blnCustomMenu        = oRs("orgCustomMenu")
		sDefaultPhone        = oRs("defaultphone")
		sDefaultEmail        = oRs("defaultemail")
		sDefaultState        = oRs("defaultstate")
		sDefaultCity         = oRs("defaultcity")
		sDefaultZip          = oRs("defaultzip")
		blnSeparateIndex     = oRs("separate_index_catalog")
		sWaiverText	         = oRs("orgwaivertext")
		sGoogleAnalyticAccnt = oRs("OrgGoogleAnalyticAccnt")
	Else
		' The Org could not be found due to a bad URL so take a shot at another URL before you crash - SJL - 1/5/2007
		
		' Close things before you leave
		oRs.Close
		Set oRs = Nothing 

		' Take a guess at what the URL should be and redirect them there.
		sCurrent = "http://www.egovlink.com/" & GetVirtualDirectyName()
		response.redirect sCurrent
	End If

	oRs.Close
	Set oRs = Nothing 

	If Not IsNull(iOrgID) Then 
		iReturnValue = iOrgID
	End If

	' RETURN VALUE
	SetOrganizationParameters = iReturnValue
	
End Function

'------------------------------------------------------------------------------------------------------------
' string GetVirtualDirectyName()
'------------------------------------------------------------------------------------------------------------
Function GetVirtualDirectyName()
	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 1) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = replace(sReturnValue,"/","")

End Function
%>
