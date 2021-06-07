<%@Codepage = 65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="action_line_global_functions.asp" //-->
<%
 Dim oActionOrg
 Set oActionOrg = New classOrganization

 LogThePage

dim blnIsMobileSub
blnIsMobileSub = false


blnFileUpload = false

'if iorgid = 37 or iorgid = 5 then
Set requestform = Server.CreateObject("Scripting.Dictionary")

	if instr(request.ServerVariables("CONTENT_TYPE"), "multipart") > 0 then
		blnFileUpload = True
		'response.write "THIS IS A FILE UPLOAD <br />"
		'response.flush
		set objUpload = Server.CreateObject("Dundas.Upload.2")
	
		'The "MaxFileSize" is set so high because if the file size is greater than this then the script bombs.
		objUpload.MaxFileSize = (31457280) ' MAX SIZE OF UPLOAD SPECIFIED IN BYTES, APPX. 30MB
		objUpload.SaveToMemory
	
	
		'response.write "<h1>" & objUpload.Form("fieldtype")(1) & "</h1>"
 		for each item in objUpload.Form
			name = item & ""
			'response.write name & " = " & objUpload.Form(name) & "<br />"
			value = ""
			for x = 0 to 100
				on error resume next
				value = value & objUpload.Form(name)(x) & ","
				on error goto 0
			next
			if right(value,1) = "," then value = left(value, len(value)-1)
			'response.write name & " = " & value & "<br />"

			session("dundasform-" & name) = value
			requestform.Add name, value

 		next
		'response.end
		'response.write "<hr>"
		
		'for each key in requestform.Keys
			'response.write key & " = " & requestform(key) & "<br />"
			'response.flush
		'next
		if requestform.Count < 1 then response.redirect "action.asp"

 		'response.end
	else
		if request.servervariables("remote_addr") = "10.4.8.13" then
				blnIsMobileSub = true
			if request.form("skip_twfusercookieid") <> "0" and isnumeric(request.form("skip_twfusercookieid")) then
				response.cookies("userid") = request.form("skip_twfusercookieid")
			end if
		end if
		'response.write "THIS IS A NORMAL FORM <br />"
		for each item in request.form
			requestform.Add item, request.form(item)
		next
		'for each key in requestform.Keys
			'response.write key & " = " & requestform(key) & "<br />"
			''response.flush
		'next
 		'response.end


	end if
 	
'end if

'if iorgid = 37 then
	'Set oXMLHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP.9.0")
'end if


'Check for valid form POST with data
'Is this a form POST?
 if request.servervariables("REQUEST_METHOD") <> "POST" then
  	'Is there data?
   	if requestform.Count < 1 then
     		response.redirect("action.asp")  'Return to the Action Form List page
  	 end if

   'Email Check - exclude @yandex.com domain for Carrborro
   	if instr(UCASE(requestform("cot_txtEmail")),"@YANDEX.COM") <> 0 then
     		response.redirect("action.asp")  'Return to the Action Form List page
    end if
 else
	'captcha
	if not blnIsMobileSub then
	strResponse = requestform("g-recaptcha-response")
	strIP = request.servervariables("REMOTE_HOST")
	strSecret = "6LcVxxwUAAAAAGGp_29X6bpiJ8YsWeNXinuUz6sx"

		Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		objWinHttp.SetTimeouts 0, 120000, 60000, 120000

		objWinHttp.Open "POST", "https://www.google.com/recaptcha/api/siteverify", False

		objWinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

		objWinHttp.Send "secret=" & strSecret & "&response=" & strResponse & "&remoteip=" & strIP


		If objWinHttp.Status = 200 Then 
			' Get the text of the response.
			transResponse = objWinHttp.ResponseText
		End If 

		' Trash our object now that we are finished with it.
		Set objWinHttp = Nothing

		if instr(transResponse, """success"": true") = 0 then
			response.redirect "action.asp"
		end if
	end if

   'Check for required fields
    for each oField in request.form
        if right(oField,4) = "/req" then
           lcl_fieldname     = ""
           lcl_value         = ""
           lcl_areacode      = ""
           lcl_exchange      = ""
           lcl_line          = ""
           lcl_streetnumber  = ""
           lcl_streetaddress = ""
           lcl_otheraddress  = ""

           lcl_fieldname = oField
           lcl_fieldname = replace(lcl_fieldname,"/req","")
           lcl_fieldname = replace(lcl_fieldname,"/number","")
           lcl_fieldname = replace(lcl_fieldname,"ef:","")
           lcl_fieldname = replace(lcl_fieldname,"-checkbox","")
           lcl_fieldname = replace(lcl_fieldname,"-radio","")
           lcl_fieldname = replace(lcl_fieldname,"-select","")
           lcl_fieldname = replace(lcl_fieldname,"-textarea","")
           lcl_fieldname = replace(lcl_fieldname,"-text","")

          'Check for a value.  Also, evaluate "special" fields (i.e. Phone = 3 fields, Issue/Problem Location, etc)
           if lcl_fieldname = "cot_txtDaytime_Phone" then
              lcl_areacode = trim(requestform("skip_user_areacode"))
              lcl_exchange = trim(requestform("skip_user_exchange"))
              lcl_line     = trim(requestform("skip_user_line"))
              lcl_value    = lcl_areacode & lcl_exchange & lcl_line

           elseif lcl_fieldname = "cot_txtFax" then
              lcl_areacode = trim(requestform("skip_fax_areacode"))
              lcl_exchange = trim(requestform("skip_fax_exchange"))
              lcl_line     = trim(requestform("skip_fax_line"))
              lcl_value    = lcl_areacode & lcl_exchange & lcl_line

           elseif lcl_fieldname = "issuelocation" then
              lcl_fieldtype = requestform(oField)

              if lcl_fieldtype = "largeaddresslist" then
                 lcl_streetnumber  = trim(requestform("residentstreetnumber"))
                 'lcl_streetaddress = trim(requestform("skip_address"))
                 lcl_streetaddress = trim(requestform("streetaddress"))
                 lcl_otheraddress  = trim(requestform("ques_issue2"))

              elseif lcl_fieldtype = "smalladdresslist" then
                 'lcl_streetaddress = trim(requestform("skip_address"))
                 lcl_streetaddress = trim(requestform("streetaddress"))
                 lcl_otheraddress  = trim(requestform("ques_issue2"))

              else
                 lcl_otheraddress  = trim(requestform("ques_issue2"))
              end if

              lcl_value = lcl_streetnumber & lcl_streetaddress & lcl_otheraddress

           else
              lcl_value = trim(requestform(lcl_fieldname))
           end if

          'If a value has not been entered into a field designated as being required then redirect back to the action line screen.
          'NOTE: "default_novalue" is used in radio/checkboxes.  It's hidden to the user.
           if (lcl_value = "" OR lcl_value = "default_novalue") and not blnIsMobileSub then

             'Send email alert
           	  lcl_alert_actionid    = requestform("actionid")
           	  lcl_alert_actiontitle = requestform("actiontitle")

              lcl_alert_subject = "Action Line ALERT (re: Javascript bypass)"

              lcl_alert_message = "<p>An attempt was made to submit a request without entering required fields.</p>" & vbcrlf
              lcl_alert_message = lcl_alert_message & "<p>" & vbcrlf
              lcl_alert_message = lcl_alert_message & "<strong>Org: </strong>" & sOrgName & " (" & iorgid & ")<br />" & vbcrlf
              lcl_alert_message = lcl_alert_message & "<strong>Form Name: </strong>" & lcl_alert_actiontitle & " (" & lcl_alert_actionid & ")<br />" & vbcrlf
              lcl_alert_message = lcl_alert_message & "<strong>Date/Time: </strong>" & now() & "<br />" & vbcrlf
              lcl_alert_message = lcl_alert_message & "<strong>First Blank Field Causing Error: </strong>[" & oField & "]" & vbcrlf
              lcl_alert_message = lcl_alert_message & "</p>" & vbcrlf

              'sendEmail "","egovsupport@eclink.com","",lcl_alert_subject,lcl_alert_message,"","Y"

              response.redirect "action.asp"
           end if

        end if
'       if left(oField,10) = "fmquestion" then
'        		sQuestionPrompt = "fmname" & replace(oField,"fmquestion","")
'        		sQuestions = sQuestions & "<b>" & requestform(sQuestionPrompt) & "</b><br>" & vbcrlf
'          sAnswer = replace(requestform(oField),"default_novalue","")
'      	end if
    next
 end if

 if requestform("frmsubjecttext") <> "" then

'	These emails are no longer wanted
'   	SendSpamFlag requestform("cot_txtEmail"), _
'                requestform("subjecttext"), _
'                 requestform("cot_txtFirst_Name"), _
'                 requestform("cot_txtLast_Name"), _
'                 clng(requestform("problemorg")), _
'                 requestform("cot_txtDaytime_Phone"), _
'                 requestform("cot_txtStreet"), _
'                 requestform("cot_txtCity"), _
'                 requestform("cot_txtState_vSlash_Province"), _
'                 requestform("cot_txtZIP_vSlash_Postal_Code"), _
'                 requestform("actiontitle"), _
'                 request.servervariables("remote_addr")

   	response.redirect("action_none.asp")
 end if
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: action_cgi.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Action Line Search Results.
'
' MODIFICATION HISTORY
' ?.?	05/08/07	Steve Loar - Changes to problem location to handle larger cities
' ?.?	08/??/07	Steve Loar - Changes to catch spam submissions.
' 3.0 03/25/08 David Boyer - Added PDF link to screen confirmation
' 3.1 04/16/08 David Boyer - Modified Issue Location to use new address fields
' 3.2 06/13/08 David Boyer - Added "Additional Text" using "Organization Features - Edit Displays"
' 3.3 11/13/08 David Boyer - Modified PDF link
' 3.4 05/29/09 David Boyer - Added check to see if "Additional Information" textarea is displayed or not.
' 3.5 06/17/09 David Boyer - Added "e=Y" to (action_respond.asp) urls in emails.
' 3.6 08/04/09 David Boyer - Added "delegate"
' 3.7 06/25/10 David Boyer - Now check a specific request to see if it is an "evalform".
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Check for org features
 lcl_orghasfeature_issue_location          = orghasfeature(iorgid,"issue location")
 lcl_orghasfeature_large_address_list      = orghasfeature(iorgid,"large address list")
 lcl_orghasfeature_hide_email_actionline   = orghasfeature(iorgid, "hide email actionline")
 lcl_orghasfeature_hide_actionline_details = orghasfeature(iorgid, "hide actionline details")
 lcl_orghasfeature_requestmergeforms       = orghasfeature(iorgid, "requestmergeforms")
 lcl_orghasfeature_actionline_usecustomfoilemailedits = orghasfeature(iorgid, "actionline_useCustomFOILEmailEdits")

'Build the page title
 lcl_page_title = "E-Gov Services " & sOrgName

 if iorgid = 7 then
    lcl_page_title = sOrgName
 end if

'Get feature name(s)
 lcl_featurename_actionline = oActionOrg.GetOrgFeatureName( "action line" )
%>
<html>
<head>
	<title><%=lcl_page_title%></title>
	<!-- This metadata is for setting the priority and importance for CDO mail messages -->
	<!--  
	METADATA  
	TYPE="typelib"  
	UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
	NAME="CDO for Windows 2000 Library"  
	--> 
	
	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />
	
	<script language="javascript" src="scripts/modules.js"></script>
	
	<script language="javascript">
	<!--
		var openWin2 = function(url, name) {
		  popupWin = window.open(url, name,"resizable,width=500,height=450");
		};

		var PDFdocument = function(p_flid,p_action_autoid,p_orgid,p_user_id) {
		   //var newFile = "http://secure.eclink.com/egovlink/action_line/actionline_pdf.asp?sys=DEV&iletterid=" + p_flid + "&action_autoid=" + p_action_autoid + "&orgid=" + p_orgid + "&userid=" + p_user_id;
		   var newFile = "viewPDF_pdf.asp?action_autoid=" + p_action_autoid + "&orgid=" + p_orgid + "&userid=" + p_user_id;
		   newWin = window.open(newFile);
		   newWin.focus();
		};
	//-->
	</script>
</head>
<!--#include file="include_top.asp"-->
<p class="title">Thank you for using the <%=sOrgName%>&nbsp;<%=lcl_featurename_actionline%>.</p>

<%
 Dim iAdminCount, aAdminUserIDs(2), aAdminEmails(2), sAdminEmail, i, lcl_resolved_status, sHideIssueLocAddInfo, lcl_group_id, blnIssueDisplay

 iAdminCount               = 0
 lcl_resolved_status       = ""
 lcl_public_actionline_pdf = ""

 adminEmailAddr   = ""
 aAdminEmails(0)  = ""
 aAdminEmails(1)  = ""
 aAdminEmails(2)  = ""
 adminid          = 0
 aAdminUserIDs(0) = ""
 aAdminUserIDs(1) = ""
 aAdminUserIDs(2) = ""
 blnIssueDisplay = false 

'PARSE TO GET TITLE OF ACTION FORM
 If instr(requestform("actionid"),"|") > 0 then
	arrForm     = split(requestform("actionid"),"|")
	actionid    = CLng(arrForm(0))
	actiontitle = arrForm(1)
	actiontitle = replace(actiontitle,"\"," > ")
	actiontitle = replace(actiontitle,"/"," > ")
 Else
	If IsNumeric(requestform("actionid")) Then 
		actionid    = CLng(requestform("actionid"))
		actiontitle = requestform("actiontitle")
	Else
		response.write "<font class=""error"">!There was an error processing this request. No form was found for this submission!</font>"
		response.End
	End If 
 End If

'MAKE SURE WE HAVE VALID REQUEST ID BEFORE PROCEEDING
 if actionid = "" then
   	response.write "<font class=""error"">!There was an error processing this request. No form was found for this submission!</font>"
   	response.end
 end if

'Get the internal default firstname, lastname, and email for this org
 lcl_internal_email = GetInternalDefaultEmail( iorgid )

 if lcl_internal_email = "" then
    lcl_internal_email = "Dev Support: No default email for orgid (" & iorgid & ") <devsupport@eclink.com>"
 else
    lcl_internal_email = "Internal Default Email <" & lcl_internal_email & ">"
 end if

'Get request form information
sSql = "SELECT assigned_userid, assigned_userid2, assigned_userid3, "
sSql = sSql & "action_form_resolved_status, public_actionline_pdf, "
sSql = sSql & "hideIssueLocAddInfo, deptid, customFOILEmailEdits, "
sSql = sSql & "redirectURL "
sSql = sSql & "FROM egov_action_request_forms "
sSql = sSql & "WHERE action_form_id = " & actionid

Dim oAdmin, bFound, sAdminLastName, sAdminFirstName, lcl_redirectURL
bFound = false 
set oAdmin = Server.CreateObject("ADODB.Recordset")
oAdmin.Open sSql, Application("DSN"), 3, 1

if not oAdmin.eof then
	lcl_resolved_status   = oAdmin("action_form_resolved_status")
	sHideIssueLocAddInfo  = oAdmin("hideIssueLocAddInfo")
	lcl_group_id          = oAdmin("deptid")
	if lcl_group_id = "" or isnull(lcl_group_id) then lcl_group_id = "0"
    	sCustomFOILEmailEdits = oAdmin("customFOILEmailEdits")
	lcl_redirectURL = oAdmin("redirectURL")

	'BEGIN: 1st ASSIGNED-TOP ----------------------------------------------------
	if oAdmin("assigned_userid") = "" or isNull(oAdmin("assigned_userid")) then
		adminEmailAddr = GetInternalDefaultEmail( iorgid )
		adminid = 0
	else
		bFound = getAdminNameAndEmail( oAdmin("assigned_userid"), sAdminEmail, sAdminLastName, sAdminFirstName )

		if bFound then
			if iorgid = 18 then
				'This handles Vandalia's inability to receive email from themselves
				adminFromAddr = "<noreply@eclink.com>"
			else 
				adminFromAddr = sAdminFirstName & " " & sAdminLastName & " <" & sAdminEmail & ">"    ' ASSIGNED ADMIN USER EMAIL
			end if

			if sAdminEmail = "" then
				sAdminEmail = lcl_internal_email
			else
				sAdminEmail = sAdminFirstName & " " & sAdminLastName & " <" & sAdminEmail & ">"   ' ASSIGNED ADMIN USER EMAIL
			end if

			adminFromAddr              = sAdminEmail
			adminEmailAddr             = sAdminEmail
			aAdminEmails(iAdminCount)  = sAdminEmail
			aAdminUserIDs(iAdminCount) = oAdmin("assigned_userid")
			adminid                    = oAdmin("assigned_userid")  'ASSIGNED ADMIN USER ID
		end if

	end if
	'END: 1st ASSIGNED-TOP ------------------------------------------------------

	'BEGIN: 2nd ASSIGNED-TOP ----------------------------------------------------
	if oAdmin("assigned_userid2") = "" or isNull(oAdmin("assigned_userid2")) then
		''nothing
	else
		bFound = false 
		sAdminEmail = ""
		sAdminLastName = ""
		sAdminFirstName = ""
		bFound = getAdminNameAndEmail( oAdmin("assigned_userid2"), sAdminEmail, sAdminLastName, sAdminFirstName )

		if bFound then
			iAdminCount = iAdminCount + 1

			if sAdminEmail = "" then
				sAdminEmail = lcl_internal_email
			else
				sAdminEmail = sAdminFirstName & " " & sAdminLastName & " <" & sAdminEmail & ">"   ' ASSIGNED ADMIN USER EMAIL
			end if

			adminEmailAddr             = adminEmailAddr & ", " & sAdminEmail   
			aAdminEmails(iAdminCount)  = sAdminEmail
			aAdminUserIDs(iAdminCount) = oAdmin("assigned_userid2")
		end If

	end if
	'END: 2nd ASSIGNED-TOP ------------------------------------------------------

	'BEGIN: 3rd ASSIGNED-TOP ----------------------------------------------------
	if oAdmin("assigned_userid3") = "" or isNull(oAdmin("assigned_userid3")) then
		''nothing
	else
		bFound = false 
		sAdminEmail = ""
		sAdminLastName = ""
		sAdminFirstName = ""
		bFound = getAdminNameAndEmail( oAdmin("assigned_userid3"), sAdminEmail, sAdminLastName, sAdminFirstName )

		if bFound then
			iAdminCount = iAdminCount + 1

			if sAdminEmail = "" then
				sAdminEmail = lcl_internal_email
			else
				sAdminEmail = sAdminFirstName & " " & sAdminLastName & " <" & sAdminEmail & ">"   ' ASSIGNED ADMIN USER EMAIL
			end if

			adminEmailAddr             = adminEmailAddr & ", " & sAdminEmail   
			aAdminEmails(iAdminCount)  = sAdminEmail
			aAdminUserIDs(iAdminCount) = oAdmin("assigned_userid3")
		end If

	end if
	'END: 3rd ASSIGNED-TOP ------------------------------------------------------

	'Get the pdf url
	lcl_public_actionline_pdf = oAdmin("public_actionline_pdf")
		
 end if

 oAdmin.Close
 set oAdmin = Nothing


'this catches the case where an admin has been assigned but they do not have and email entered for them
'If Trim(adminEmailAddr) = "" Then 
'' if trim(sAdminEmail) = "" then
''	adminEmailAddr   = "<" & GetInternalDefaultEmail( iorgid ) & ">"
''	aAdminEmails(0)  = adminEmailAddr
''	aAdminEmails(1)  = adminEmailAddr
''	aAdminEmails(2)  = adminEmailAddr
''	adminid          = 0
''	aAdminUserIDs(0) = ""
''	aAdminUserIDs(1) = ""
''	aAdminUserIDs(2) = ""
 'end if

'GET QUESTIONS AND ENTERED VALUES INFORMATION
' This creates the blob'
sQuestions = ""

for x = 1 to 50
	sQuestionPrompt = "fmname" & x
	if requestform.Exists(sQuestionPrompt) then
		'response.write "fmname" & x & ": " & requestform("fmname")
		lcl_question    = requestform(sQuestionPrompt)
		lcl_question    = replace(lcl_question,"&quot;","""")

		sQuestions = sQuestions & "<p>" & vbcrlf
		sQuestions = sQuestions & "<b>" & lcl_question & "</b><br>" & vbcrlf

		'sAnswer = replace(requestform(oField),"default_novalue","")
		sAnswer = replace(requestform("fmquestion" & x),"default_novalue","")
		sAnswer = replace(sAnswer,"&quot;","""")

		'if trim(sAnswer) <> "" then
			sFormattedAnswer = sAnswer
		'end if

		sQuestions = sQuestions & dbsafe(sFormattedAnswer) & vbcrlf
		sQuestions = sQuestions & "</p>" & vbcrlf
	else
		exit for
	end if
next
'response.end


sQuestions2 = sQuestions

Dim iCitizenUserId

'Get User Information
if sOrgRegistration then
	if request.cookies("userid") <> "" and request.cookies("userid") <> "-1" then
		iCitizenUserId = request.cookies("userid")
		
		'They are logged in, so update their user info
        	updateCitizenInfo iCitizenUserId, _
			requestform("cot_txtFirst_Name"), _
			requestform("cot_txtLast_Name"), _
			requestform("cot_txtEmail"), _
			requestform("cot_txtBusiness_Name"), _
			requestform("cot_txtDaytime_Phone"), _
			requestform("cot_txtFax"), _
			requestform("cot_txtStreet"), _
			requestform("cot_txtCity"), _
			requestform("cot_txtState_vSlash_Province"), _
			requestform("cot_txtZip_vSlash_Postal_Code"), _
			requestform("cot_txtCountry")
	else 
		' They are not logged in, so create a new user account
		iCitizenUserId = AddUserInformation()
	end if
else
 	' The city does not have user registration, so always create new user accounts
	iCitizenUserId = AddUserInformation()
end if

' Create the action line request in the system
sSql = "INSERT INTO egov_actionline_requests ( userid, orgid, assignedemployeeid, comment, "
sSql = sSql & "category_id, category_title, groupid, status, submittedby_remoteaddress, submit_date,mobileoption_latitude,mobileoption_longitude,isfrommobile,  "
sSql = sSql & "media_url, mobileappdescription, service_request_id ) VALUES ( "
sSql = sSql & iCitizenUserId & ", " 
sSql = sSql & iorgid & ", "
if adminid <> 0 then 
	sSql = sSql & adminid & ", "
Else
	sSql = sSql & " NULL, "
End If 
sSql = sSql & "'" & replace(sQuestions,"'","''") & "', "
sSql = sSql & actionid & ", "
sSql = sSql & "'" & dbsafe( actiontitle ) & "', "
sSql = sSql & lcl_group_id & ", "
sSql = sSql & "'SUBMITTED', "
sSql = sSql & "'" & dbsafe( Request.ServerVariables("REMOTE_ADDR") ) & "', "
datOrgDateTime = ConvertDateTimetoTimeZone(iorgid)

sMyLatitude = "NULL"
sMyLongitude = "NULL"
if requestform("myLatitude") <> "" then sMyLatitude = "'" & dbsafe(requestform("myLatitude")) & "'"
if requestform("myLongitude") <> "" then sMyLongitude = "'" & dbsafe(requestform("myLongitude")) & "'"
if requestform("mapLat") <> "" then sMyLatitude = "'" & dbsafe(requestform("mapLat")) & "'"
if requestform("mapLng") <> "" then sMyLongitude = "'" & dbsafe(requestform("mapLng")) & "'"
sIsFromMobile = "0"
if requestform("isfrommobile") <> "" then sIsFromMobile = "1"


sSql = sSql & "'" & datOrgDateTime & "', " & sMyLatitude & ", " & sMyLongitude & ", " & sIsFromMobile & ", "

media_url = "NULL"
if requestform("media_url") <> "" then media_url = "'" & dbsafe(requestform("media_url")) & "'"
mobileappdescription = "NULL"
if requestform("mobileappdescription") <> "" then 
	mobileappdescription = "'" & dbsafe(requestform("mobileappdescription")) & "'"
	sQuestions = mobileappdescription
end if
service_request_id = "NULL"
if requestform("service_request_id") <> "" then service_request_id = "'" & dbsafe(requestform("service_request_id")) & "'"


sSql = sSql & media_url & ", " & mobileappdescription & ", " & service_request_id & ") "
'response.write sSql & "<br><br>"

iTrackingNumber = RunIdentityInsertStatement( sSql )


if iorgid = 153 and actionid = 17051 then
	tmpDate = datOrgDateTime
	DueDate = GetFOILDueDate(tmpDate)


	sSQL = "UPDATE egov_actionline_requests SET due_date = '" & DueDate & "' WHERE action_autoid = " & iTrackingNumber
	RunSQLStatement sSQL

end if


'Determine if this is an "evalform" then set the status to "EVALFORM"
 lcl_status     = "RESOLVED"

 lcl_isEvalForm = False
 lcl_isEvalForm = checkOrgForm(iorgid, "EvaluationFormID", actionid)

 if lcl_isEvalForm then
    lcl_status          = "EVALFORM"
    lcl_resolved_status = "Y"
 end if

'If the status is RESOLVED then populate the Activity Request Log so that the user can see that it has been resolved.
 if lcl_resolved_status = "Y" then

	'First update the status on the Activity Request
	sSql = "UPDATE egov_actionline_requests SET "
	sSql = sSql & " status = '" & lcl_status & "', "
	sSql = sSql & " complete_date = '" & datOrgDateTime & "'"
	sSql = sSql & " WHERE action_autoid = " & iTrackingNumber
	
	RunSQLStatement sSql

	AddCommentTaskComment "Request's status was set to " & lcl_status & " upon submission.", _
		"This request was submitted by " & requestform("cot_txtFirst_Name") & " " & requestform("cot_txtLast_Name") & ".", _
		lcl_status, iTrackingNumber, lcl_user_id, iorgid, datOrgDateTime

 end if

'REPLACES BLOB FUNCTIONALITY - STORES DATA IN PROMPT ANSWER FORMAT
InsertRequestFieldsandResponses iTrackingNumber

'GENERATE TRACKING NUMBER - (FORMULA IS SQL ROWID + HHMM)
lngTrackingNumber = iTrackingNumber & replace(FormatDateTime(datOrgDateTime,4),":","")


'Setup the problem location address
' This is only used for sending emails, not for storing in the DB
lcl_problem_location = ""

if lcl_orghasfeature_issue_location then

  	'Retrieve the problem/location title for this form.
   	sSql = "SELECT ISNULL(issuelocationname,'ISSUE/PROBLEM LOCATION:') AS issuelocationname, action_form_display_issue "
   	sSql = sSql & "FROM egov_action_request_forms WHERE action_form_id = " & actionid

   	Set oForm = Server.CreateObject("ADODB.Recordset")
   	oForm.Open sSql, Application("DSN"), 0, 1

   	If Not oForm.EOF Then
  	   	sIssueName      = UCASE(oForm("issuelocationname"))
  	   	If Trim(sIssueName) = "" Then
   		    	sIssueName = "ISSUE/PROBLEM LOCATION:"
     		End If
     		blnIssueDisplay = oForm("action_form_display_issue")
     	Else
     		blnIssueDisplay = false 
     	End If
     	
     	oForm.Close
     	set oForm = Nothing 

	'1. Check to see if the "issue location" feature has been "turned on" for this form.
	'2. Check to see if the org has the large address list feature "turned on"
	'3. Check to see if a street number has been entered.
	if blnIssueDisplay = True then
	   	
	   	if lcl_orghasfeature_large_address_list then
	       	lcl_problem_location = requestform("residentstreetnumber")

			'4. Check to see if a value in the dropdown list has been selected.
			'It doesn't matter if the large address feature has been turned on/off.
			'If an org has the "issue location" feature then the dropdown will appear.
	 		if requestform("streetaddress") <> "0000" then
				if lcl_problem_location <> "" then
					lcl_problem_location = lcl_problem_location & " " & requestform("streetaddress")
				else
					lcl_problem_location = requestform("streetaddress")
				end if
	            else
				'5. If no value has been selected in the dropdown list then check
				'to see if the "other" address has been populated.
				'If it has then override the street number, if it was populated.
				'If not then display whatever has been entered.  The screen will enforce
				'a value to be entered for the address if the street number has been entered.
		            if requestform("ques_issue2") <> "" then
					lcl_problem_location = requestform("ques_issue2")
				end if
	 		end if
		else
			'6. If the org does NOT have the "large address list" feature "turned on"
			'then if a value has been selected from the dropdown list retrieve the street address
			if requestform("streetaddress") <> "0000" then
				sSql = "SELECT residentstreetnumber, residentstreetname "
				sSql = sSql & " FROM egov_residentaddresses "
				sSql = sSql & " WHERE residentaddressid = " & requestform("streetaddress")

				set rs = Server.CreateObject("ADODB.Recordset")
				rs.Open sSql, Application("DSN"), 0, 1

				if not rs.eof then
					if rs("residentstreetnumber") <> "" then
						lcl_problem_location = rs("residentstreetnumber")
					end if

					if rs("residentstreetname") <> "" then
						if lcl_problem_location <> "" then
							lcl_problem_location = lcl_problem_location & " " & rs("residentstreetname")
						else
							lcl_problem_location = rs("residentstreetname")
						end if
					end if
				end if
				
				rs.Close
				Set rs = Nothing 
	 		else
				if requestform("ques_issue2") <> "" then
	   				lcl_problem_location = requestform("ques_issue2")
	  			end if
 			end if
 		end if  'END orghasfeature("large address list")
	end if  'END blnIssueDisplay

 end if  'END orghasfeature("issue location")
 
 
'Record Location Information in the DB
' This is after the above code that get issue location info for the emails so we can use 
' the blnIssueDisplay it finds and not have to do this inside the AddIssueLocation function again
AddIssueLocation iTrackingNumber, blnIssueDisplay


'BEGIN: Build message and send email to citizen ------------------------------
if iorgid <> "7" Then
	lcl_customFOILmsg = ""

	if lcl_orghasfeature_actionline_usecustomfoilemailedits then
		lcl_customFOILmsg = sCustomFOILEmailEdits
		lcl_customFOILmsg = "&nbsp;&nbsp;" & lcl_customFOILmsg
	end if

	sMsg = sMsg & "<p>This automated message was sent by the " & sOrgName & " " & lcl_featurename_actionline & ".  "

	if Not lcl_orghasfeature_actionline_usecustomfoilemailedits then
		sMsg = sMsg & "Do not reply to this message.  "
		sMsg = sMsg & "Follow the instructions below " & vbcrlf

		if not lcl_orghasfeature_hide_email_actionline then
			sMsg = sMsg & "or contact " & adminFromAddr 
		end If

		sMsg = sMsg & " for inquiries regarding this email.</p>" & vbcrlf 
		sMsg = sMsg & " " & vbcrlf
	end if

	sMsg = sMsg & "<p>Thank you for submitting your information to " & sOrgName & " on " & datOrgDateTime & "." & lcl_customFOILmsg & "</p>" & vbcrlf 
	sMsg = sMsg & " " & vbcrlf 

	if requestform("actionid") <> "17890" then
		sMsg = sMsg & "<p>To check the status, please follow this link to the " & sOrgName & " web site at...<br />"  & vbcrlf 
		sMsg = sMsg & " <a href=""" & sEgovWebsiteURL & "/action_request_lookup.asp?request_id=" & lngTrackingNumber & """>" & sEgovWebsiteURL & "/action_request_lookup.asp?request_id=" & lngTrackingNumber & "</a><br />" & vbcrlf 
		sMsg = sMsg & "Make sure that the entire URL appears in your browser's address field.</p>" & vbcrlf 
		sMsg = sMsg & " " & vbcrlf
	end if

	sMsg = sMsg & "<p><strong>TRACKING NUMBER: " & lngTrackingNumber & " </strong></p>" & vbcrlf
	'sMsg = sMsg & "CATEGORY: " & UCASE(actiontitle)   & vbcrlf
	if requestform("actionid") <> "17890" then
	sMsg = sMsg & "<p><strong>CATEGORY:</strong> " & actiontitle & "</p>" & vbcrlf
	sMsg = sMsg & "<p><strong>SUGGESTION/ISSUE:</strong>  </p>" & vbcrlf 
	end if
	sMsg = sMsg & "<p>" & vbcrlf & replace(sQuestions,"default_novalue","") & vbcrlf & "</p>"
	'sMsg = sMsg & "<p>" & vbcrlf & sQuestions & vbcrlf & "</p>"

	if lcl_orghasfeature_issue_location then
		if blnIssueDisplay then
			sMsg = sMsg & "<p><strong>" & sIssueName & "</strong><br />" & vbcrlf
			sMsg = sMsg & "LOCATION: " & lcl_problem_location & "<br />" & vbcrlf

			if not sHideIssueLocAddInfo then
				sMsg = sMsg & "ADDITIONAL INFORMATION: " & requestform("ques_issue6") & "<br />" & vbcrlf
			end if
		end if
	end if

	'sMsg = sMsg & "<p>" & vbcrlf & vbcrlf  & fnPlainText(sQuestions) & vbcrlf & vbcrlf & "</p>"
	if requestform("actionid") <> "17890" then
	sMsg = sMsg & "<p>We will evaluate this request and take action as appropriate.</p>" & vbcrlf 
	sMsg = sMsg & " " & vbcrlf 
	else
		sMsg = sMsg & "<p><b><a href=""http://www.egovlink.com/rye/rye_rock_memo.asp?id=" & iTrackingNumber & """>Print Registration for Posting on Property</a></b></p>"
		sMsg = sMsg & " " & vbcrlf 
		sMsg = sMsg & "<p><b><a href=""http://www.ryeny.gov/rockremovalreg.cfm"">View Registration List</a></b></p>"
		sMsg = sMsg & " " & vbcrlf 
	end if
	sMsg = sMsg & "<p>Thank you for using the " & sOrgName & " " & lcl_featurename_actionline & ".  We want to understand what you want and expect, make it easier for you to do business with us, and to respond as quickly as practical to your requests.</p>" & vbcrlf 
	sMsg = sMsg & " " & vbcrlf 

	lcl_subject = "Submission: " & actiontitle & " (re: " & lngTrackingNumber & ")"
	lcl_message = BuildHTMLMessage(sMsg)
else
	lcl_subject = "Submission: EC Link HelpDesk - HelpDesk Ticket"
	lcl_message = BuildHTMLMessage(BuildHTMLBody())
end if


'Determine if the "Email Confirmation" checkbox was "checked"
if requestform("chkSendEmail") = "YES" then

	'If an email for the contact (citizen) exists then validate the format of the email and send the email.
	if requestform("cot_txtEmail") <> "" then

		'Remove the name from the email address
		lcl_validate_email = formatSendToEmail(requestform("cot_txtEmail"))

		if isValidEmail(lcl_validate_email) then

			'Send the email
			sendEmail "",requestform("cot_txtEmail"),"",lcl_subject,lcl_message,"","Y"
			if requestform("actionid") = "17890" then
				'sendEmail "","rockreg@ryeny.gov;tfoster@eclink.com","",lcl_subject,lcl_message,"","Y"
				sendEmail "","rockreg@ryeny.gov","",lcl_subject,lcl_message,"","Y"
			end if
		else
			ErrorCode = 1
		end if

		'Add to email queue if unsuccessful
		if ErrorCode <> 0 then
			sMsg      = Left(sMsg,5000)
			SendToAdd = requestform("cot_txtEmail")
			fnPlaceEmailinQueue Application("SMTP_Server"),sOrgName & " E-GOV WEBSITE",adminFromAddr,SendToAdd,sOrgName & " E-GOV MSG - " & UCase(lcl_featurename_actionline) & " REQUEST",1,sMsg,1,-1
		end if
	end if
end if
			
if ErrorCode <> 0 then
	'ADD LOGGING CODE HERE
	response.write "The request has been logged but there was an error sending an email notice to you.  You will not receive an email notice.<br /><br /><br />" & vbcrlf
	'response.write ErrorCode & "<br>"
	'response.write Err.Number  & "<br>"
	'response.write Err.Description  & "<br>"
	bMailSent1 = False
end if
'END: Build message and send email to citizen --------------------------------


'BEGIN: Build message and send email to site administrator -------------------
if iorgid <> "7" then
	sMsg2 = sMsg2 & "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.  Contact " & adminFromAddr & " for inquiries regarding this email.</p>" & vbcrlf
	sMsg2 = sMsg2 & " " & vbcrlf
	sMsg2 = sMsg2 & "<p>A " & sOrgName & " " & lcl_featurename_actionline & " issue was submitted on " & datOrgDateTime & ".</p>" & vbcrlf
	sMsg2 = sMsg2 & " " & vbcrlf
	sMsg2 = sMsg2 & "<p><strong>Click the following link to view this " & lcl_featurename_actionline & " Request:</strong><br />" & vbcrlf
	sMsg2 = sMsg2 & "<a href=""" & sEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & iTrackingNumber & "&e=Y"">" & vbcrlf
	sMsg2 = sMsg2 & sEgovWebsiteURL & "/admin/action_line/action_respond.asp?control=" & iTrackingNumber & "&e=Y</a></p>" & vbcrlf

	if not lcl_orghasfeature_hide_actionline_details then
		sMsg2 = sMsg2 & " " & vbcrlf
		sMsg2 = sMsg2 & "<p><strong>" & UCase(lcl_featurename_actionline) & " REQUEST DETAILS</strong><br />" & vbcrlf
		sMsg2 = sMsg2 & "DATE SUBMITTED: "  & datOrgDateTime    & "<br />" & vbcrlf 
		sMsg2 = sMsg2 & "TRACKING NUMBER: " & lngTrackingNumber & "<br />" & vbcrlf 
		sMsg2 = sMsg2 & "CATEGORY ID: "     & actionid          & "<br />" & vbcrlf 
		sMsg2 = sMsg2 & "CATEGORY Title: "  & actiontitle       & "</p>"   & vbcrlf 
		sMsg2 = sMsg2 & "<p><strong>SUGGESTION/ISSUE: ...</strong><br />"  & vbcrlf
		'sMsg2 = sMsg2 & "" & vbcrlf & vbcrlf  & REPLACE(sQuestions,"default_novalue","") & vbcrlf & vbcrlf & "</p>"
		sMsg2 = sMsg2 & "" & vbcrlf & sQuestions & vbcrlf & "</p>" & vbcrlf

		if lcl_orghasfeature_issue_location then
			if blnIssueDisplay then
				sMsg2 = sMsg2 & "<p><strong>" & sIssueName & "</strong><br />" & vbcrlf
				sMsg2 = sMsg2 & "LOCATION: "               & lcl_problem_location   & "<br />" & vbcrlf

				if not sHideIssueLocAddInfo then
					sMsg2 = sMsg2 & "ADDITIONAL INFORMATION: " & requestform("ques_issue6") & "<br />" & vbcrlf
				end if

			end if
		end if

		sMsg2 = sMsg2 & "<p><strong>" & UCase(lcl_featurename_actionline) & " REQUESTER CONTACT INFORMATION</strong><br />" & vbcrlf
		sMsg2 = sMsg2 & "NAME: " & requestform("cot_txtFirst_Name") & " " & requestform("cot_txtLast_Name") & "<br />" & vbcrlf
		sMsg2 = sMsg2 & "BUSINESS: " & requestform("cot_txtBusiness_Name") & "<br />" & vbcrlf
		'sMsg2 = sMsg2 & "<br />EMAIL: " & requestform("cot_txtEmail") & vbcrlf

		if requestform("cot_txtEmail") <> "" then
			sMsg2 = sMsg2 & "EMAIL: <a href=mailto:" & requestform("cot_txtEmail") &">" & requestform("cot_txtEmail") & "</a>" & "<br />" & vbcrlf
		end if

		sMsg2 = sMsg2 & "PHONE: "   & FormatPhone(requestform("cot_txtDaytime_Phone")) & "<br />" & vbcrlf
		sMsg2 = sMsg2 & "FAX: "     & FormatPhone(requestform("cot_txtFax"))           & "<br />" & vbcrlf
		sMsg2 = sMsg2 & "ADDRESS: " & requestform("cot_txtStreet")                     & "<br />" & vbcrlf
		sMsg2 = sMsg2 & requestform("cot_txtCity") & " " & requestform("cot_txtState_vSlash_Province") & "<br />" & vbcrlf
		sMsg2 = sMsg2 & requestform("cot_txtZIP_vSlash_Postal_Code") & " " & requestform("cot_txtCountry") & vbcrlf & "</p>" & vbcrlf
	end if

	lcl_subject = "Submission: " & actiontitle & " (re: " & requestform("cot_txtFirst_Name") & " " & requestform("cot_txtLast_Name") & ")"
	lcl_message = BuildHTMLMessage( sMsg2 )
else
	'lcl_from    = sOrgName & " ECLink HelpDesk <webmaster@eclink.com>"
	lcl_from    = sOrgName & " ECLink HelpDesk <noreply@eclink.com>"
	lcl_subject = "Submission: EC Link HelpDesk - HelpDesk Ticket"
	lcl_message = BuildHTMLMessage(BuildAdminHTMLBody())

	'Send the email (to the HelpDesk)
	sendEmail lcl_from, adminEmailAddr,"",lcl_subject,lcl_message,"","Y"
end if

'Cycle through the admin emails and send to admin(s) only if an email exist and is valid.
for iEmailCount = 0 to iAdminCount
	'Remove the name from the email address
	lcl_validate_email = formatSendToEmail(aAdminEmails(iEmailCount))

	if isValidEmail(lcl_validate_email) then

		'Check for a delegate
		getDelegateInfo aAdminUserIDs(iEmailCount), lcl_delegateid, lcl_delegate_username, lcl_delegate_useremail

		'Setup the SENDTO and check for a DELEGATE
		setupSendToAndDelegateEmails aAdminEmails(iEmailCount), lcl_delegate_useremail, lcl_email_sendto, lcl_email_cc

		'Send the email
		sendEmail "",lcl_email_sendto,lcl_email_cc,lcl_subject,lcl_message,"","Y"
	end if
next
'END: Build message and send email to site administrator ---------------------


'UPLOAD THE FILE IF IT EXISTS
if blnFileUpload then
	for each objFile in objUpload.Files
		lcl_uploadfile = objFile.OriginalPath

    		strFileName = LCASE(RIGHT(lcl_uploadfile,LEN(lcl_uploadfile) - instrrev(lcl_uploadfile,"\")))
	
        	strBasePath = "E:\egovlink300_docs\custom\pub\" & GetVirtualDirectyName() & "\"
    		checkFolder strBasePath 
		strBasePath = strBasePath & "mobile_uploads\"
    		checkFolder strBasePath 
		strBasePath = strBasePath & lngTrackingNumber & "\"
    		checkFolder strBasePath 
	
		'Store the file in the server filesystem
		if objUpload.FileExists( strBasePath & "\" & strFileName ) then
			'Delete the file if it already exists on server filesystem
			objUpload.FileDelete( strBasePath & strFileName )
		end if
	
		'Save file on server filesystem
		objFile.SaveAs(  strBasePath & strFileName )
	next


end if

set objUpload = Nothing



'REDIRECT?
if lcl_redirectURL <> "" and not blnIsMobileSub then
	response.redirect lcl_redirectURL & "?tn=" & lngTrackingNumber
end if
'response.end



'BEGIN: Display Information to the user --------------------------------------
response.write "<div class=""box_header2"">" & vbcrlf

if OrgHasDisplay(iorgid,"action submitted title" ) then
	response.write GetOrgDisplay( iOrgId, "action submitted title" )
else
	response.write lcl_featurename_actionline & " Request Submitted"
end if

response.write " - " & datOrgDateTime &  "</div>" & vbcrlf
response.write "  <div class=""groupSmall"">" & vbcrlf
if requestform("actionid") = "17890" then
	response.write "    Thank you for your registration.<p>" & vbcrlf
else
response.write "    A request has been submitted under the subject of <b><i>" & actiontitle & "</i></b>.  " & vbcrlf
response.write "    We'll evaluate your request and take action as appropriate.<p>" & vbcrlf
end if

response.write "    Your tracking number is <b>" & lngTrackingNumber & "</b>. Please record this number for your records.<p>" & vbcrlf

if OrgHasDisplay(iorgid,"actionline_publicconfirm_checkstatusofrequest" ) AND GetOrgDisplay(iOrgID, "actionline_publicconfirm_checkstatusofrequest") <> "" then
	response.write "    <p>" & vbcrlf
	response.write      GetOrgDisplay(iOrgID, "actionline_publicconfirm_checkstatusofrequest")
	response.write "    </p>" & vbcrlf
elseif  requestform("actionid") <> "17890" then
	response.write "    <p>" & vbcrlf
	response.write "    You can check the status of the request by visiting the <b>" & lcl_featurename_actionline & " </b> main page. " & vbcrlf
	response.write "    Simply enter the above tracking number to review the status of the request. " & vbcrlf
	response.write "    Please allow at least 24 hours for a response to your request." & vbcrlf
	response.write "    </p>" & vbcrlf
end if

if requestform("actionid") = "17890" then
	response.write "    <p>" & vbcrlf
	response.write "    <b><a href=""rye_rock_memo.asp?id=" & iTrackingNumber & """>Print Registration for Posting on Property</a></b>" & vbcrlf
	response.write "    </p>" & vbcrlf
end if

'Determine if there is "additional text" set up in "Organization Features - Edit Displays
if OrgHasDisplay( iorgid, "actionline_confirmtext_public") then
	lcl_display_id      = 0
	lcl_additional_text = ""

	lcl_display_id      = GetDisplayId("actionline_confirmtext_public")
	lcl_additional_text = GetOrgDisplayWithId(iorgid,lcl_display_id,true)

	'Display Additional Text if a value exists
	if lcl_additional_text <> "" then
		response.write "    <p>" & vbcrlf
		response.write      lcl_additional_text & vbcrlf
		response.write "    </p>" & vbcrlf
	end if
end if

'Determine if the "View PDF" button is displayed.
' 1. The org must be assigned the "requestmergeforms" feature
' 2. The form on the request has a PDF associated to it.
if lcl_orghasfeature_requestmergeforms AND lcl_public_actionline_pdf <> "" then
	response.write "<!--center>" & vbcrlf
	response.write "<input type=""button"" class=""button"" onClick=""window.open('viewXMLPDF.asp?iRequestID=" & iTrackingNumber & "');"" value=""View Request in PDF Format"" />" & vbcrlf
	response.write "</center-->" & vbcrlf
end if

response.write "  </div>"
response.write "</div>" & vbcrlf

Set oActionOrg = Nothing 

%>

<!--#Include file="include_bottom.asp"-->  

<%
'------------------------------------------------------------------------------
Function AddUserInformation()
	Dim sSql 
	
	sSql = "INSERT INTO egov_users ( userfname, userlname, useremail, userbusinessname, "
	sSql = sSql & "userhomephone, userfax, useraddress, usercity, userstate, userzip, usercountry ) VALUES ("
	sSql = sSql & " '" & dbsafe(requestform("cot_txtFirst_Name")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtLast_Name")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtEmail")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtBusiness_Name")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtDaytime_Phone")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtFax")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtStreet")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtCity")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtState_vSlash_Province")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtZIP_vSlash_Postal_Code")) & "',"
	sSql = sSql & " '" & dbsafe(requestform("cot_txtCountry")) & "'"
	sSql = sSql & " )"
	'if iorgid = 113 then
		'response.write sSql & "<br><br>"
		'response.write requestform("cot_txtFirst_Name")
		'response.end
	'end if
	
	AddUserInformation = RunIdentityInsertStatement( sSql )
	
End Function

'------------------------------------------------------------------------------
sub updateCitizenInfo( ByVal iUserID, ByVal iUserFName, ByVal iUserLName, ByVal iUserEmail, ByVal iUserBusinessName, ByVal iUserHomePhone, ByVal iUserFax, ByVal iUserAddress, ByVal iUserCity, ByVal iUserState, ByVal iUserZip, ByVal iUserCountry )
	Dim sSql, sUserID, sUserFName, sUserLName, sUserEmail, sUserBusinessName
	Dim sUserHomePhone, sUserFax, sUserAddress, sUserCity, sUserState, sUserZip, sUserCountry

	sUserID           = 0
	sUserFName        = ""
	sUserLName        = ""
	sUserEmail        = ""
	sUserBusinessName = ""
	sUserHomePhone    = ""
	sUserFax          = ""
	sUserAddress      = ""
	sUserCity         = ""
	sUserState        = ""
	sUserZip          = ""
	sUserCountry      = ""

	if iUserID <> "" then
		if not containsApostrophe(iUserID) then
			sUserID = CLng(iUserID)
		end if
	end if

	if iUserFName <> "" then
		sUserFName = iUserFName
	end if

	if iUserLName <> "" then
		sUserLName = iUserLName
	end if

	if iUserEmail <> "" then
		sUserEmail = iUserEmail
	end if

	if iUserBusinessName <> "" then
		sUserBusinessName = iUserBusinessName
	end if

	if iUserHomePhone <> "" then
		sUserHomePhone = iUserHomePhone
	end if

	if iUserFax <> "" then
		sUserFax = iUserFax
	end if

	if iUserAddress <> "" then
		sUserAddress = iUserAddress
	end if

	if iUserCity <> "" then
		sUserCity = iUserCity
	end if

	if iUserState <> "" then
		sUserState = iUserState
	end if

	if iUserZip <> "" then
		sUserZip = iUserZip
	end if

	if iUserCountry <> "" then
		sUserCountry = iUserCountry
	end if

	if sUserID > 0 then
		if sUserFName <> "" then
			sUserFName = dbsafe(sUserFName)
			sUserFName = "'" & sUserFName & "'"

			sSql = " userfname = " & sUserFName
		end if

		if sUserLName <> "" then
			sUserLName = dbsafe(sUserLName)
			sUserLName = "'" & sUserLName & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " userlname = " & sUserLName
		end if

		if sUserEmail <> "" then
			sUserEmail = dbsafe(sUserEmail)
			sUserEmail = "'" & sUserEmail & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " useremail = " & sUserEmail
		end if

		if sUserBusinessName <> "" then
			sUserBusinessName = dbsafe(sUserBusinessName)
			sUserBusinessName = "'" & sUserBusinessName & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " userbusinessname = " & sUserBusinessName
		end if

		if sUserHomePhone <> "" then
			sUserHomePhone = dbsafe(sUserHomePhone)
			sUserHomePhone = "'" & sUserHomePhone & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " userhomephone = " & sUserHomePhone
		end if

		if sUserFax <> "" then
			sUserFax = dbsafe(sUserFax)
			sUserFax = "'" & sUserFax & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " userfax = " & sUserFax
		end if

		if sUserAddress <> "" then
			sUserAddress = dbsafe(sUserAddress)
			sUserAddress = "'" & sUserAddress & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " useraddress = " & sUserAddress
		end if

		if sUserCity <> "" then
			sUserCity = dbsafe(sUserCity)
			sUserCity = "'" & sUserCity & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " usercity = " & sUserCity
		end if

		if sUserState <> "" then
			sUserState = dbsafe(sUserState)
			sUserState = "'" & sUserState & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " userstate = " & sUserState
		end if

		if sUserZip <> "" then
			sUserZip = dbsafe(sUserZip)
			sUserZip = "'" & sUserZip & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " userzip = " & sUserZip
		end if

		if sUserCountry <> "" then
			sUserCountry = dbsafe(sUserCountry)
			sUserCountry = "'" & sUserCountry & "'"

			if sSql <> "" then
				sSql = sSql & ", "
			end if

			sSql = sSql & " usercountry = " & sUserCountry
		end if

		If sSql <> "" Then 
			sSql = " UPDATE egov_users SET " & sSql
			sSql = sSql & " WHERE userid = " & sUserID

			RunSQLStatement sSql
		End If 
	end if

end sub


'------------------------------------------------------------------------------
Function fnPlaceEmailinQueue( ByVal sHost, ByVal sFromName, ByVal sFromEmail, ByVal sSendEmail, ByVal sSubject, ByVal iBodyFormat, ByVal sBodyMessage, ByVal iPriority, ByVal iErrorCode)

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = "AddEmailtoFailoverQueue"
		.CommandType = adCmdStoredProc
		.Parameters.Append oCmd.CreateParameter("Host", adVarChar , adParamInput, 50, sHost)
		.Parameters.Append oCmd.CreateParameter("FromName", adVarChar, adParamInput, 50, sFromName)
		.Parameters.Append oCmd.CreateParameter("FromEmail", adVarChar, adParamInput, 255, sFromEmail)
		.Parameters.Append oCmd.CreateParameter("SendEmail", adVarChar, adParamInput, 255, sSendEmail)
		.Parameters.Append oCmd.CreateParameter("Subject", adVarChar, adParamInput, 1024, sSubject)
		.Parameters.Append oCmd.CreateParameter("BodyFormat", adInteger, adParamInput, 4, iBodyFormat)
		.Parameters.Append oCmd.CreateParameter("BodyMessage", adVarChar, adParamInput, 5000, sBodyMessage)
		.Parameters.Append oCmd.CreateParameter("Priority", adInteger, adParamInput, 4, iPriority)
		.Parameters.Append oCmd.CreateParameter("ErrorCode", adVarChar, adParamInput, 10, iErrorCode)
		.Execute
	End With

	Set oCmd = Nothing
	
End Function

'------------------------------------------------------------------------------
Function fnPlainText( ByVal sValue)
	'sValue = UCASE(sValue)  Removed per Peter on 3/14/2006 - Steve Loar
	sValue = replace(sValue,"<B>","")
	sValue = replace(sValue,"</B>","")
	sValue = replace(sValue,"<P>","")
	sValue = replace(sValue,"</P>",vbcrlf)
	sValue = replace(sValue,"<BR>",vbcrlf)
	sValue = replace(sValue,"</BR>",vbcrlf)
	sValue = replace(sValue,"<STRONG>","")
	sValue = replace(sValue,"</STRONG>","")

	sValue = replace(sValue,"<b>","")
	sValue = replace(sValue,"</b>","")
	sValue = replace(sValue,"<p>","")
	sValue = replace(sValue,"</p>",vbcrlf)
	sValue = replace(sValue,"<br>",vbcrlf)
	sValue = replace(sValue,"</br>",vbcrlf)
	sValue = replace(sValue,"<br />",vbcrlf)
	sValue = replace(sValue,"<strong>","")
	sValue = replace(sValue,"</strong>","")

	fnPlainText = sValue

End Function

'------------------------------------------------------------------------------
Function BuildHTMLBody()
	Dim sMsgNew

	sMsgNew = sMsgNew & "This automated message was sent by the ECLINK HELPDESK web site. Do not reply to this message.  Please follow the instructions below or contact <strong>" & adminFromAddr & "</strong> for inquiries regarding this email.<br />"
	sMsgNew = sMsgNew & "<br /> " & vbcrlf 
	sMsgNew = sMsgNew & "You created a helpdesk ticket on " & datOrgDateTime & ".<br />" & vbcrlf 
	sMsgNew = sMsgNew & "<br />" & vbcrlf 
	sMsgNew = sMsgNew & "Tracking Number: <b> " 
	sMsgNew = sMsgNew & lngTrackingNumber & "</b><br /> " & vbcrlf 
	sMsgNew = sMsgNew & "<br /><br /> " & vbcrlf 
	sMsgNew = sMsgNew & "<strong>HELP DESK FORM:</strong> " & UCASE(actiontitle)  & "<br /><br />" & vbcrlf & vbcrlf
	sMsgNew = sMsgNew & "<strong>HELP DESK TICKET DETAILS:</strong><br /><br />"
	sMsgNew = sMsgNew & vbcrlf & sQuestions2 & vbcrlf
	sMsgNew = sMsgNew & "<br /><br /><br />We will evaluate this issue and take action as appropriate.<br />" & vbcrlf 
	sMsgNew = sMsgNew & "<br /> " & vbcrlf 
	sMsgNew = sMsgNew & "Thank you for using our help desk web site to better serve your business needs.<br />" & vbcrlf 
	sMsgNew = sMsgNew & "<br /> " & vbcrlf 

	BuildHTMLBody = sMsgNew

End Function

'------------------------------------------------------------------------------
Function BuildAdminHTMLBody()
	Dim sMsgAdmin

	sMsgAdmin = sMsgAdmin & "This automated message was sent by the ECLINK HELPDESK web site. Do not reply to this message.  Follow the instructions below or contact <strong>" & adminFromAddr & "</strong> for inquiries regarding this email.<br />" 
	sMsgAdmin = sMsgAdmin & "<br /> " & vbcrlf 
	sMsgAdmin = sMsgAdmin & "To follow-up with this help desk ticket please follow the link below:<br /><br />http://www.egovlink.com/" & sorgVirtualSiteName & "/admin<br /><br />" & vbcrlf 
	sMsgAdmin = sMsgAdmin & "<br /><strong>DATE SUBMITTED:</strong> " & datOrgDateTime & vbcrlf
	sMsgAdmin = sMsgAdmin & "<br /><strong>TRACKING NUMBER:</strong> " & lngTrackingNumber & vbcrlf
	sMsgAdmin = sMsgAdmin & "<br /><strong>HELP DESK FORM:</strong> " & actiontitle & vbcrlf
	sMsgAdmin = sMsgAdmin & "<br /><strong>HELP DESK TICKET DETAILS:</strong><br /><br />" 
	sMsgAdmin = sMsgAdmin & vbcrlf & sQuestions2 & vbcrlf
	sMsgAdmin = sMsgAdmin & "<br /><br /><strong>TICKET SUBMITTER CONTACT INFORMATION</strong>" & vbcrlf
	sMsgAdmin = sMsgAdmin & "<br />NAME: " & requestform("cot_txtFirst_Name") & " " & requestform("cot_txtLast_Name") & vbcrlf 
	sMsgAdmin = sMsgAdmin & "<br />BUSINESS: " & requestform("cot_txtBusiness_Name") & vbcrlf 
	sMsgAdmin = sMsgAdmin & "<br />EMAIL: " & requestform("cot_txtEmail") & vbcrlf 
	sMsgAdmin = sMsgAdmin & "<br />PHONE: " & requestform("cot_txtDaytime_Phone") & vbcrlf 
	sMsgAdmin = sMsgAdmin & "<br />FAX: " & requestform("cot_txtFax") & vbcrlf 
	sMsgAdmin = sMsgAdmin & "<br />ADDRESS: " & requestform("cot_txtStreet") & vbcrlf
	sMsgAdmin = sMsgAdmin & "<br />" & requestform("cot_txtCity") & " " & requestform("cot_txtState_vSlash_Province") & " " 
	sMsgAdmin = sMsgAdmin & "<br />" & requestform("cot_txtZIP_vSlash_Postal_Code") & " " & requestform("cot_txtCountry") & vbcrlf & vbcrlf

	BuildAdminHTMLBody = sMsgAdmin

End Function


'----------------------------------------------------------------------------------------
Sub AddIssueLocation( ByVal iTrackingNumber, blnIssueDisplay )
	Dim iReturnValue, oLocation, sNumber, sStreetPrefix, sStreetName, sStreetSuffix, sStreetDirection, sStreetUnit, sCity, sState, sZip
	Dim sLatitude, sLongitude, sValidStreet, sCounty, sParcelID, sExcludeFromAL
	
	iReturnValue      = 0
	sNumber           = ""
	sStreetPrefix     = ""
	sStreetName       = ""
	sStreetSuffix     = ""
	sStreetDirection  = ""
	sStreetUnit       = ""
	sSortStreetName   = ""
	sCity             = oActionOrg.GetDefaultCity()
	sState            = oActionOrg.GetDefaultState()
	sZip              = oActionOrg.GetDefaultZip()
	sLatitude         = "0.00"
	sLongitude        = "0.00"
	sValidStreet      = requestform("validstreet") ' This is either a Y or N'
	sCounty           = ""
	sParcelID         = ""
	sExcludeFromAL    = 0
	sListedOwner      = ""
	sResidentType     = "N"
	sLegalDesc        = ""
	sRegisteredUserID = 0

	if requestform("streetaddress") <> "" then
		if lcl_orghasfeature_large_address_list then

			if requestform("streetaddress") <> "0000" then
				'Try to match the input street number and selected street name to those in the database
				MatchAddressInfo requestform("residentstreetnumber"), requestform("streetaddress"), _
					sNumber, sPrefix, sStreetName, sSuffix, sDirection, sCity, sState, sZip, sLatitude, _
					sLongitude, sCounty, sParcelID, sExcludeFromAL, sListedOwner, sResidentType, sLegalDesc, sRegisteredUserID
			else
				'they selected Other address not listed
				BreakOutAddress requestform("ques_issue2"), sNumber, sStreetName
			end if

		else
			' They do not have large addresses
			if CLng(requestform("streetaddress")) <> CLng(0) then
				'Handle the dropdown addresses - These should have the residentaddressid as the selected value
				GetAddressInfo requestform("streetaddress"), sNumber, sPrefix, _
					sStreetName, sSuffix, sDirection, sCity, sState, sZip, sLatitude, sLongitude, _
					sCounty, sParcelID, sExcludeFromAL, sListedOwner, sResidentType, sLegalDesc, sRegisteredUserID
				
				sValidStreet = "Y"
			else
				'they selected Other address not listed
				BreakOutAddress requestform("ques_issue2"), sNumber, sStreetName
			end if
		end if

	else
		sValidStreet = "Y"
	end if

	'Build the SortStreetName -----------------------------------------------------
	sSortStreetName = trim(sStreetName)

	if trim(sSuffix) <> "" then
		if sSortStreetName <> "" then
			sSortStreetName = sSortStreetName & " " & sSuffix
		else
			sSortStreetName = sSuffix
		end if
	end if

	if trim(sDirection) <> "" then
		if sSortStreetName <> "" then
			sSortStreetName = sSortStreetName & " " & sDirection
		else
			sSortStreetName = sDirection
		end if
	end if

	if trim(sPrefix) <> "" then
		if sSortStreetName <> "" then
			sSortStreetName = sSortStreetName & " " & sPrefix
		else
			sSortStreetName = sPrefix
		end if
	end if


	'If the form has the issue location and the org has the issue location feature and the field is blank then default the value
	If blnIssueDisplay And (sStreetName = "" Or ISNULL(sStreetName)) Then 
		sStreetName = "Street Address has not been entered."
	End If 
	
	If CDbl(sLatitude) = CDbl("0.00") then
		sLatitude = "NULL"
		sLongitude = "NULL"
	End If 
	
	If sExcludeFromAL Then
		sExcludeFromAL = 1
	Else 
		sExcludeFromAL = 0
	End If 

	' Do the insert Here. 
	sSql = "INSERT INTO egov_action_response_issue_location ( actionrequestresponseid, streetnumber, streetprefix, streetaddress, streetsuffix, "
	sSql = sSql & "streetdirection, city, state, zip, validstreet, county, parcelidnumber, excludefromactionline, listedowner, "
	sSql = sSql & "legaldescription, residenttype, registereduserid, sortstreetname, latitude, longitude, streetunit, comments ) VALUES ( "
	sSql = sSql & iTrackingNumber & ", "
	sSql = sSql & "'" & sNumber & "', "
	sSql = sSql & "'" & dbsafe(sPrefix) & "', "
	sSql = sSql & "'" & dbsafe(sStreetName) & "', "
	sSql = sSql & "'" & dbsafe(sSuffix) & "', "
	sSql = sSql & "'" & dbsafe(sDirection) & "', "
	sSql = sSql & "'" & dbsafe(sCity) & "', "
	sSql = sSql & "'" & dbsafe(sState) & "', "
	sSql = sSql & "'" & dbsafe(sZip) & "', "
	sSql = sSql & "'" & sValidStreet & "', "
	sSql = sSql & "'" & dbsafe(sCounty) & "', "
	sSql = sSql & "'" & dbsafe(sParcelID) & "', "
	sSql = sSql & sExcludeFromAL & ", "
	sSql = sSql & "'" & dbsafe(sListedOwner) & "', "
	sSql = sSql & "'" & dbsafe(sLegalDesc) & "', "
	sSql = sSql & "'" & sResidentType & "', "
	sSql = sSql & sRegisteredUserID & ", "
	sSql = sSql & "'" & dbsafe(sSortStreetName) & "', "
	sSql = sSql & sLatitude & ", "
	sSql = sSql & sLongitude & ", "
	sSql = sSql & "'" & dbsafe(requestform("streetunit")) & "', "
	sSql = sSql & "'" & dbsafe(requestform("ques_issue6")) & "' )"
	'response.write sSql & "<br><br>"
		
	RunSQLStatement sSql 

end sub


'------------------------------------------------------------------------------
sub MatchAddressInfo( ByVal sResidentStreetNumber, ByVal sResidentStreetName, ByRef sNumber, ByRef sPrefix, ByRef sStreetName, _
                      ByRef sSuffix, ByRef sDirection, ByRef sCity, ByRef sState, ByRef sZip, ByRef sLatitude, ByRef sLongitude, ByRef sCounty, _
                      ByRef sParcelID, ByRef sExcludeFromAL, ByRef sListedOwner, ByRef sResidentType, ByRef sLegalDesc, ByRef sRegisteredUserID )
	dim sSql, oAddress

	sSql = "SELECT residentstreetnumber, "
	sSql = sSql & " residentstreetprefix, "
	sSql = sSql & " residentstreetname, "
	sSql = sSql & " streetsuffix, "
	sSql = sSql & " streetdirection, "
	sSql = sSql & " residentcity, "
	sSql = sSql & " residentstate, "
	sSql = sSql & " residentzip, "
	sSql = sSql & " isnull(latitude,0.00) as latitude, "
	sSql = sSql & " isnull(longitude,0.00) as longitude, "
	sSql = sSql & " county, "
	sSql = sSql & " parcelidnumber, "
	sSql = sSql & " excludefromactionline, "
	sSql = sSql & " listedowner, "
	sSql = sSql & " residenttype, "
	sSql = sSql & " legaldescription, "
	sSql = sSql & " isnull(registereduserid,0) as registereduserid "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & iorgid
	sSql = sSql & " AND excludefromactionline = 0 "
	sSql = sSql & " AND UPPER(residentstreetnumber) = UPPER('" & dbsafe(sResidentStreetNumber) & "') "
	sSql = sSql & " AND (residentstreetname = '" & dbsafe(sResidentStreetName) & "' "
	sSql = sSql & " OR residentstreetname + ' ' + streetsuffix = '" & dbsafe(sResidentStreetName) & "' "
	sSql = sSql & " OR residentstreetname + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "' "
	sSql = sSql & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "' "
	sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname = '" & dbsafe(sResidentStreetName) & "' "
	sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & dbsafe(sResidentStreetName) & "' "
	sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "' "
	sSql = sSql & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & dbsafe(sResidentStreetName) & "'"
	sSql = sSql & ")"

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSql, Application("DSN"), 3, 1

	if Not oAddress.eof then
		sNumber           = oAddress("residentstreetnumber")
		sPrefix           = oAddress("residentstreetprefix")
		sStreetName       = oAddress("residentstreetname")
		sSuffix           = oAddress("streetsuffix")
		sDirection        = oAddress("streetdirection")
		sCity             = oAddress("residentcity")
		sState            = oAddress("residentstate")
		sZip              = oAddress("residentzip")
		sLatitude         = oAddress("latitude")
		sLongitude        = oAddress("longitude")
		sCounty           = oAddress("county")
		sParcelID         = oAddress("parcelidnumber")
		sExcludeFromAL    = oAddress("excludefromactionline")
		sListedOwner      = oAddress("listedowner")
		sResidentType     = oAddress("residenttype")
		sLegalDesc        = oAddress("legaldescription")
		sRegisteredUserID = oAddress("registereduserid")
	else
		if sResidentStreetNumber <> "" then
			'sStreetName = dbsafe(sResidentStreetNumber) & " " & dbsafe(sResidentStreetName)
			sStreetName = sResidentStreetNumber & " " & sResidentStreetName
		else
			'sStreetName = dbsafe(sResidentStreetName)
			sStreetName = sResidentStreetName
		end if
	end if

	oAddress.Close
	set oAddress = nothing

end sub


'------------------------------------------------------------------------------
sub GetAddressInfo( ByVal sResidentAddressId, ByRef sNumber, ByRef sPrefix, ByRef sStreetName, ByRef sSuffix, ByRef sDirection, _
                    ByRef sCity, ByRef sState, ByRef sZip, ByRef sLatitude, ByRef sLongitude, ByRef sCounty, ByRef sParcelID, _
                    ByRef sExcludeFromAl, ByRef sListedOwner, ByRef sResidentType, ByRef sLegalDesc, ByRef sRegisteredUserID )
	dim sSql, oAddress

	sSql = "SELECT residentstreetnumber, "
	sSql = sSql & " residentstreetprefix, "
	sSql = sSql & " residentstreetname, "
	sSql = sSql & " streetsuffix, "
	sSql = sSql & " streetdirection, "
	sSql = sSql & " residentcity, "
	sSql = sSql & " residentstate, "
	sSql = sSql & " residentzip, "
	sSql = sSql & " isnull(latitude,0.00) as latitude, "
	sSql = sSql & " isnull(longitude,0.00) as longitude, "
	sSql = sSql & " county, "
	sSql = sSql & " parcelidnumber, "
	sSql = sSql & " excludefromactionline, "
	sSql = sSql & " listedowner, "
	sSql = sSql & " residenttype, "
	sSql = sSql & " legaldescription, "
	sSql = sSql & " isnull(registereduserid,0) as registereduserid "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE residentaddressid = " & sResidentAddressId 
	sSql = sSql & " AND excludefromactionline = 0 "

	set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSql, Application("DSN"), 3, 1

	if Not oAddress.EOF then
		sNumber           = oAddress("residentstreetnumber")
		sPrefix           = oAddress("residentstreetprefix")
		sStreetName       = oAddress("residentstreetname")
		sSuffix           = oAddress("streetsuffix")
		sDirection        = oAddress("streetdirection")
		sCity             = oAddress("residentcity")
		sState            = oAddress("residentstate")
		sZip              = oAddress("residentzip")
		sLatitude         = oAddress("latitude")
		sLongitude        = oAddress("longitude")
		sCounty           = oAddress("county")
		sParcelID         = oAddress("parcelidnumber")
		sExcludeFromAL    = oAddress("excludefromactionline")
		sListedOwner      = oAddress("listedowner")
		sResidentType     = oAddress("residenttype")
		sLegalDesc        = oAddress("legaldescription")
		sRegisteredUserID = oAddress("registereduserid")
	end if

	oAddress.close
	set oAddress = nothing

end sub


'------------------------------------------------------------------------------
Function FormatPhone( ByVal Number )

	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If
  
End Function


'------------------------------------------------------------------------------
sub InsertRequestFieldsandResponses( ByVal iRequestID )
	Dim iFieldCount

	iFieldCount = 0

	'Enumerate fields and entered responses
	for each oField in requestform.Keys

		'Get only fields and their associated value
		if len(oField) >=10 then
			if left(oField,10) = "fmquestion" then

				'Get the Field Prompt
				iFieldCount = iFieldCount + 1
				'response.write iFieldCount & "<br />"
				'response.write requestform("fieldtype") & "<br />"
				'response.flush
				strFieldType = (split(replace(requestform("fieldtype")," ",""),","))(iFieldCount-1)
				on error resume next
				strAnswerList = (split(replace(requestform("answerlist")," ",""),","))(iFieldCount-1)
				on error goto 0
				strIsRequired = (split(replace(requestform("isRequired")," ",""),","))(iFieldCount-1)
				strSequence = (split(replace(requestform("sequence")," ",""),","))(iFieldCount-1)
				strPushFieldID = (split(replace(requestform("pushfieldid")," ",""),","))(iFieldCount-1)
				strFmName = requestform("fmname" & replace(oField,"fmquestion",""))
				iFieldID    = InsertFieldPrompt(strFmName, iRequestID, strFieldType, strAnswerList, strIsRequired, strSequence, strPushFieldID)
				'response.write strFmName & ":" & iFieldID & "<br />"
				
				'Enumerate and get field responses
				intLoopCount = UBOUND(split(replace(requestform(oField)," ",""),","))
				if intLoopCount = 0 then intLoopCount = 1
				'response.write oField & " UBOUND:" & intLoopCount & "<br />"
				'response.flush


				if intLoopCount > 1 then
					blnExitNow = false
					for iResponse = 1 to intLoopCount
						strIResponse = ""
						on error resume next
						strIResponse = cstr(requestform(oField)(iResponse-1)) & ""
						on error goto 0
						if strIResponse = "" then 
							strIResponse = requestform(oField) & ""
							intLoopCount = 1
							blnExitNow = true
						end if
						'response.write strIResponse & " HERE<br />"
						strPDFFormName = ""
						on error resume next
						strPDFFormName = requestform("pdfformname")(iFieldCount-1)
						on error goto 0
	
					'if iorgid = 37 then
						'response.write "HERE1" & strIResponse & "," & iFieldID & "," & strPDFFormName & "," & strPushFieldId & " HERE<br />"
						'response.end
					'end if
						InsertFieldResponse strIResponse, iFieldID, strPDFFormName, strPushFieldId
						if blnExitNow then Exit For
					next
				else
					strIResponse = requestform(oField) & ""
					'if iorgid = 37 then
						'response.write "HERE2" & strIResponse & "," & iFieldID & "," & strPDFFormName & "," & strPushFieldId & " HERE<br />"
						'response.end
					'end if
					on error resume next
					strPDFFormName = requestform("pdfformname")(iFieldCount-1)
					on error goto 0
					InsertFieldResponse strIResponse, iFieldID, strPDFFormName, strPushFieldId
				end if
				'response.write "<hr>"
	
			end if
		end if
	next

end sub


'------------------------------------------------------------------------------
function InsertFieldPrompt( ByVal sPrompt, ByVal iRequestID, ByVal iFieldType, ByVal sAnswerList, ByVal blnIsRequired, ByVal iSequence, ByVal iPushFieldID )
	Dim lcl_fieldprompt, sSql, iReturnId, iIsRequired
	
	iReturnId = 0

	'Format field prompt/answerlist
	lcl_fieldprompt = sPrompt
	lcl_fieldprompt = replace(sPrompt,"&quot;","""")

	lcl_answerlist  = sAnswerList
	lcl_answerlist  = replace(lcl_answerlist,"&quot;","""")

	if trim(iSequence) = "" then
		iSequence = "NULL"
	end if

	if trim(iPushFieldID) = "" then
		lcl_pushfieldid = 0
	else
		lcl_pushfieldid = trim(iPushFieldID)
	end if
	
	If blnIsRequired Then 
		iIsRequired = 1
	Else
		iIsRequired = 0
	End If

	sSql = "INSERT INTO egov_submitted_request_fields ( submitted_request_field_prompt, submitted_request_field_type_id, " 
	sSql = sSql & "submitted_request_field_answerlist, submitted_request_field_isrequired, submitted_request_field_pdf_name, "
	sSql = sSql & "submitted_request_field_pushfieldid, submitted_request_field_sequence, submitted_request_id ) VALUES ( "
	sSql = sSql & "'" & dbsafe(lcl_fieldprompt) & "', "
	sSql = sSql & iFieldType & ", "
	sSql = sSql & "'" & dbsafe(lcl_answerlist) & "', "
	sSql = sSql & iIsRequired & ", "
	sSql = sSql & "'" & dbsafe(sPDF_Name) & "', "
	sSql = sSql & lcl_pushfieldid & ", "
	sSql = sSql & iSequence & ", "
	sSql = sSql & iRequestID & " )"
	'response.write "Question: " & sSql & "<br><br>"

	iReturnId = RunIdentityInsertStatement( sSql )
	
	InsertFieldPrompt = iReturnId

end function

'------------------------------------------------------------------------------
Sub InsertFieldResponse( ByVal sResponse, ByVal iFieldID, ByVal sPDFName, ByVal sPushFieldID )
	Dim lcl_response, sSql, lcl_pushfieldid

	if sResponse <> "" then
		lcl_response = sResponse
		lcl_response = replace(lcl_response,"&quot;","""")
	else
		lcl_response = ""
	end if

	if trim(sPushFieldID) = "" then
		lcl_pushfieldid = 0
	else
		lcl_pushfieldid = trim(sPushFieldID)
	end if
	
	sSql = "INSERT INTO egov_submitted_request_field_responses ( submitted_request_field_id, "
	sSql = sSql & "submitted_request_field_response, submitted_request_form_field_name, "
	sSql = sSql & "submitted_request_pushfieldid ) VALUES ( "
	sSql = sSql & iFieldID & ", "
	sSql = sSql & "'" & dbsafe(lcl_response) & "', "
	sSql = sSql & "'" & dbsafe(sPDFName) & "', "
	sSql = sSql & lcl_pushfieldid & " )"
	'response.write "Response: " & sSql & "<br><br>"
		
	RunSQLStatement sSql

End Sub

'------------------------------------------------------------------------------
Sub SendSpamFlag( ByVal sFromEmail, ByVal sTextinput, ByVal sFirstName, ByVal sLastName, ByVal iOrgId, _
                 ByVal sPhone, ByVal sStreet, ByVal sCity, ByVal sState, ByVal sZip, ByVal sFormTitle, ByVal sIPAddress )

	Dim oCdoMail, oCdoConfm, sMsgBody, sOrgName, sQuestions, sFormattedAnswer, sQuestionPrompt, sAnswer, iResponse

	sOrgName = GetOrgname( iOrgId )

	sMsgBody = "Possible spam submitted to " & sOrgName   & ".<br />" & vbcrlf
	sMsgBody = sMsgBody & "Form Name: "      & sFormTitle & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Citizen Name: "   & sFirstName & " " & sLastName & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Email Address: "  & sFromEmail & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Phone: "          & sPhone     & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Address: "        & sStreet    & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "City: "           & sCity      & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "State: "          & sState     & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Zip: "            & sZip       & "<br />" & vbcrlf
 	sMsgBody = sMsgBody & "IP Address: "     & sIPAddress & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "Hidden field contains: " & sTextinput & "<br />" & vbcrlf
	sMsgBody = sMsgBody & "<p>Other Fields contain: </p>" & vbcrlf & vbcrlf

	'Get Questions and Entered Values Information
	sQuestions = ""

	for each oField in request.form
		sAnswer          = ""
		sFormattedAnswer = ""

   		if left(oField,10) = "fmquestion" then

     			sQuestionPrompt = "fmname" & replace(oField,"fmquestion","")

				sQuestions = sQuestions & "<p>" & vbcrlf
				sQuestions = sQuestions & "<b>" & requestform(sQuestionPrompt) & "</b><br />" & vbcrlf 

     			for iResponse = 1 to requestform(oField).count
				sAnswer = sAnswer & replace(requestform(oField)(iResponse),"default_novalue","") & "<br />" & vbcrlf
     			next

			if trim(sAnswer) <> "" then
			   sFormattedAnswer = sAnswer
			end if

			sQuestions = sQuestions & sFormattedAnswer & vbcrlf
			sQuestions = sQuestions & "</p>" & vbcrlf
		end if
	next
	sMsgBody = sMsgBody & sQuestions

	'sendEmail UCase(sOrgName) & " E-GOV WEBSITE <noreply@eclink.com>", "jfelix@eclink.com, pselden@eclink.com", "", UCase(sOrgName) & " E-GOV POSSIBLE SPAM SUBMISSION", sMsgBody, "", "Y"

	Set oCdoMail = Nothing 
	Set oCdoConf = Nothing 

End Sub 

'------------------------------------------------------------------------------
Function GetOrgname( ByVal iOrgId )
	Dim sSql, oOrgInfo

	sSql = "SELECT OrgName FROM Organizations WHERE orgid = " & iOrgId 

	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSql, Application("DSN"), 3, 1
	
	If Not oOrgInfo.EOF Then
		GetOrgname = oOrgInfo("OrgName")
	Else
		GetOrgname = ""
	End If 

	oOrgInfo.Close
	Set oOrgInfo = Nothing 

End Function 


'------------------------------------------------------------------------------
Sub AddCommentTaskComment( ByVal sInternalMsg, ByVal sExternalMsg, ByVal sStatus, ByVal iTrackingNumber, ByVal iUserID, ByVal iOrgID, ByVal sCurrentDate )
	Dim sSql

	sSql = "INSERT egov_action_responses ("
	sSql = sSql & "action_status,"
	sSql = sSql & "action_internalcomment,"
	sSql = sSql & "action_externalcomment,"
	sSql = sSql & "action_editdate,"
	sSql = sSql & "action_userid,"
	sSql = sSql & "action_orgid,"
	sSql = sSql & "action_autoid"
	sSql = sSql & ") VALUES ("
	sSql = sSql & "'" & sStatus              & "', "
	sSql = sSql & "'" & DBsafe(sInternalMsg) & "', "
	sSql = sSql & "'" & DBsafe(sExternalMsg) & "', "
	sSql = sSql & "'" & sCurrentDate         & "', "
	sSql = sSql & "'" & iUserID              & "',"
	sSql = sSql & "'" & iOrgID               & "',"
	sSql = sSql & "'" & iTrackingNumber      & "')"

	RunSQLStatement sSql 

End Sub


'------------------------------------------------------------------------------
Function GetTimeOffset( ByVal iOrgID )
	Dim sSql, oRs

	sSql = "SELECT T.gmtoffset "
	sSql = sSql & " FROM Organizations O, TimeZones T "
	sSql = sSql & " WHERE O.OrgTimeZoneID = T.TimeZoneID "
	sSql = sSql & " AND O.orgid = " & iOrgID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetTimeOffset =  clng(oRs("gmtoffset"))
	Else
		GetTimeOffset = clng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
' boolean = getAdminNameAndEmail( iUserId, sEmail, sLastName, sFirstName )
'------------------------------------------------------------------------------
Function getAdminNameAndEmail( ByVal iUserId, ByRef sEmail, ByRef sLastName, ByRef sFirstName )
	Dim sSql, oRs, bFound
	
	sSql = "SELECT ISNULL(email,'') AS email, ISNULL(lastname,'') AS lastname, ISNULL(firstname,'') AS firstname "
	sSql = sSql & "FROM users WHERE UserId = " & iUserId 
	'response.write sSql & "<br><br>"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		sEmail = oRs("email")
		sLastName = oRs("lastname")
		sFirstName = oRs("firstname")
		bFound = true 
	Else
		sEmail = ""
		sLastName = ""
		sFirstName = ""
		bFound = false 
	End If 

	oRs.Close
	Set oRs = Nothing 
	
	getAdminNameAndEmail = bFound

End Function


'------------------------------------------------------------------------------
function dbsafe( ByVal p_value )
	Dim lcl_value
	
	lcl_value = ""

	If Not VarType( p_value ) = vbString Then dbsafe = p_value : Exit Function

	lcl_value = replace(p_value,"'","''")
	lcl_value = replace(lcl_value, "<", "&lt;" )

	dbsafe = lcl_value

end function


'------------------------------------------------------------------------------
sub dtb_debug( ByVal p_value )
	Dim sSql, oDTB
	
	sSql = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
	set oDTB = Server.CreateObject("ADODB.Recordset")
	oDTB.Open sSql, Application("DSN"), 3, 1

	set oDTB = nothing

end sub


%>
<!--#include file="../egovlink300_global/includes/inc_rye.asp"-->
<%

'------------------------------------------------------------------------------
Sub LogThePage( )
	Dim sSql, oCmd, sScriptName, sVirtualDirectory, aVirtualDirectory, sPage, arr, sUserAgent, sUserAgentGroup

	sScriptName = Request.ServerVariables("SCRIPT_NAME")

	If request.servervariables("http_user_agent") <> "" Then 
		sUserAgent = "'" & Track_DBsafe(Trim(Left(request.servervariables("http_user_agent"),480))) & "'"
	Else
		sUserAgent = "NULL"
	End If 

	If Len(Trim(request.servervariables("http_user_agent"))) > 0 Then 
		sUserAgentGroup = "'" & GetUserAgentGroup( LCase(request.servervariables("http_user_agent")) ) & "'"
	Else
		sUserAgentGroup = "'" & GetUntrackedUserAgentGroup( ) & "'"
	End If 

	' Get the virtual directory
	aVirtualDirectory = Split(sScriptName, "/", -1, 1) 
	sVirtualDirectory = "/" & aVirtualDirectory(1) 
	sVirtualDirectory = "'" & Replace(sVirtualDirectory,"/","") & "'"

	' Get the page
	For Each arr in aVirtualDirectory 
		sPage = arr 
	Next 

	sSql = "INSERT INTO egov_pagelog ( virtualdirectory, applicationside, page, loadtime, scriptname, querystring, "
	sSql = sSql & " servername, remoteaddress, requestmethod, orgid, userid, username, sectionid, documenttitle, useragent, useragentgroup, requestformcollection, cookiescollection, sessioncollection, sessionid  ) VALUES ( "
	sSql = sSql & sVirtualDirectory & ", "
	sSql = sSql & "'public', "
	sSql = sSql & "'" & sPage & "', "
	sSql = sSql & FormatNumber(iLoadTime,3,,,0) & ", "
	sSql = sSql & "'" & sScriptName & "', "

	If Request.ServerVariables("QUERY_STRING") <> "" Then 
		sSql = sSql & "'" & Track_DBsafe(Left(Request.ServerVariables("QUERY_STRING"),500)) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 
	' our server name
	sSql = sSql & "'" & Request.ServerVariables("SERVER_NAME") & "', "

	' remote address
	sSql = sSql & "'" & Request.ServerVariables("REMOTE_ADDR") & "', "

	' request method - GET or POST
	sSql = sSql & "'" & Request.ServerVariables("REQUEST_METHOD") & "', "

	' orgid
	If iorgid <> "" Then 
		sSql = sSql & iorgid & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Userid
	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" and isnumeric(request.cookies("userid")) Then
		sSql = sSql & request.cookies("userid") & ", "
	Else
		sSql = sSql & "NULL, "
		response.cookies("userid") = ""
	End If 

	' Get username
	If sUserName <> "" Then
		sSql = sSql & "'" & Track_DBsafe(sUserName) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Section Id for the old LogPageVisit functionality
	If iSectionID <> "" Then 
		sSql = sSql & iSectionID & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Document Title for the old LogPageVisit functionality
	If sDocumentTitle <> "" Then 
		sSql = sSql & "'" & Track_DBsafe(sDocumentTitle) & "',  "
	Else
		sSql = sSql & "NULL, "
	End If 

	' User Agent
	sSql = sSql & sUserAgent & ", "

	' User Agent Group
	sSql = sSql & sUserAgentGroup & ", "

	sSql = sSql & "'" & Track_DBsafe(GetRequestformInformation()) & "',"
	sSql = sSql & "'" & GetCookiesCollection() & "',"
	sSql = sSql & "'" & GetSessionCollection() & "',"


	sSql = sSql & "'" & Session.SessionID & "'"

	sSql = sSql & " )"
	'response.write sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql

	session("sSql") = sSql
	oCmd.Execute
	session("sSql") = ""

	Set oCmd = Nothing


End Sub 

'------------------------------------------------------------------------------
sub checkFolder(sFolderPath)

	set oFSO = server.createobject("Scripting.FileSystemObject")
	
	if oFSO.FolderExists(sFolderPath) <> True then
  		set oFolder = oFSO.CreateFolder(sFolderPath)
		set oFolder = nothing
 	end if

	set oFSO = nothing

end sub
%>
