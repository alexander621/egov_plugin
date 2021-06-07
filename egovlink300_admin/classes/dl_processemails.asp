<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: dl_processemails.asp
' AUTHOR: Steve Loar
' CREATED: 07/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This sends the Subscriptions, Jobs and Bids emails and is called via AJAX by dl_sendmail.asp
'
' MODIFICATION HISTORY
' 1.0   07/28/2009  Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	Dim lcl_list_type, lcl_dl_logid

	lcl_list_type = request("listtype")

	' Save what is to be emailed and to whom
	lcl_dl_logid = saveSubscriptionInfo( lcl_list_type )

	' Send out the emails
	subProcessEmail

	' Update the log
	updateSubscriptionInfoStatus lcl_dl_logid, "COMPLETED"

	response.write "COMPLETED"

'-------------------------------------------------------------------------------------------------
' Function saveSubscriptionInfo( ByVal iListType )
'-------------------------------------------------------------------------------------------------
Function saveSubscriptionInfo( ByVal iListType )
	Dim sSql, sListType, lcl_sentdate, lcl_completedate, lcl_sendstatus, lcl_email_fromname
	Dim lcl_email_fromemail, lcl_email_subject, lcl_email_body, sdllist, lcl_containsHTML
	Dim iLength

	'Validate the fields
	sListType           = "NULL"
	
	' we want this to be in their timeframe
	lcl_sentdate        = " dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
	'lcl_sentdate        = "'" & Now & "'"

	lcl_completedate    = "NULL"
	lcl_sendstatus      = "'INPROGRESS'"
	lcl_email_fromname  = "NULL"
	lcl_email_fromemail = "NULL"
	lcl_email_subject   = "NULL"
	lcl_email_body      = "NULL"
	lcl_email_format    = "NULL"
	sdllist             = "NULL"
	lcl_containsHTML    = 0

	If iListType <> "" Then 
		sListType = "'" & UCase(dbsafe(iListType)) & "'"
	End If 

	If request("sFromName") <> "" Then 
		lcl_email_fromname = "'" & dbsafe(request("sFromName")) & "'"
	End If 

	If request("sFromEmail") <> "" Then 
		lcl_email_fromemail = "'" & dbsafe(request("sFromEmail")) & "'"
	End If 

	If request("sSubjectLine") <> "" Then 
		lcl_email_subject = "'" & dbsafe(request("sSubjectLine")) & "'"
	End If 

dtb_debug("BEFORE: " & request("sHTMLBody"))
	If request("sHTMLBody") <> "" Then 
    lcl_email_body = URLDecode(request("sHTMLBody"))
  		lcl_email_body = replace(lcl_email_body, Chr(10), Chr(13)+Chr(10))
'    lcl_email_body = replace(lcl_email_body, "’", "'")
  		lcl_email_body = "'" & dbsafe(lcl_email_body) & "'"
	End If 
dtb_debug("AFTER: " & lcl_email_body)

	If request("iEmailFormat") <> "" Then 
		lcl_email_format = "'" & dbsafe(request("iEmailFormat")) & "'"
	End If 

	'Get the distribution list(s)
	If request("MailList") <> "" Then 
		sdllist = replace(request("Maillist"),"X",",")  'Build comma separate list
		iLength = LEN(sdllist)
		sdllist = LEFT(sdllist,(iLength - 1))  'Trim trailing comma
		sdllist = "'" & dbsafe(sdllist) & "'"
	End If 

	If request("containsHTML") = "Y" Then 
		lcl_containsHTML = 1
	End If 

	sSql = "INSERT INTO egov_class_distributionlist_log ( "
	sSql = sSql & "orgid, "
	sSql = sSql & "distributionlisttype, "
	sSql = sSql & "sentbyuserid, "
	sSql = sSql & "sentdate, "
	sSql = sSql & "completedate, "
	sSql = sSql & "sendstatus, "
	sSql = sSql & "email_fromname, "
	sSql = sSql & "email_fromemail, "
	sSql = sSql & "email_subject, "
	sSql = sSql & "email_body, "
	sSql = sSql & "email_format, "
	sSql = sSql & "containsHTML, "
	sSql = sSql & "dl_listids "
	sSql = sSql & " ) VALUES ( "
	sSql = sSql & session("orgid")    & ", "
	sSql = sSql & sListType           & ", "
	sSql = sSql & session("userid")   & ", "
	sSql = sSql & lcl_sentdate        & ", "
	sSql = sSql & lcl_completedate    & ", "
	sSql = sSql & lcl_sendstatus      & ", "
	sSql = sSql & lcl_email_fromname  & ", "
	sSql = sSql & lcl_email_fromemail & ", "
	sSql = sSql & lcl_email_subject   & ", "
	sSql = sSql & lcl_email_body      & ", "
	sSql = sSql & lcl_email_format    & ", "
	sSql = sSql & lcl_containsHTML    & ", "
	sSql = sSql & sdllist
	sSql = sSql & ")"

'	response.write "sSql = " & sSql & "<br /><br />"

	lcl_dl_logid = RunInsertStatement(sSql)

	saveSubscriptionInfo = lcl_dl_logid

End Function 


'-------------------------------------------------------------------------------------------------
' Sub subProcessEmail( )
'-------------------------------------------------------------------------------------------------
Sub subProcessEmail( )
	Dim iRowCount, sSql, sdllist, iLength, oRs, sHTMLBody

	iRowCount = CLng(0) 

	'Get the distribution list(s)
	sdllist = Replace(request("Maillist"),"X",",") ' BUILD COMMA SEPARATE LIST
	iLength = Len(sdllist) - 1
	sdllist = Left(sdllist, iLength) ' TRIM TRAILING COMMA

	' Cleanup the escape tags put in by the encodeURIComponent Javascript function
	sHTMLBody = Replace(request("sHTMLBody"), Chr(10), Chr(13)+Chr(10))
	sHTMLBody = Replace(request("sHTMLBody"), Chr(38), "&")
	sHTMLBody = Replace(request("sHTMLBody"), Chr(59), ";")

	'GET LIST OF UNIQUE EMAIL ADDRESSES
	sSql = "SELECT DISTINCT userid, useremail "
	sSql = sSql & " FROM egov_dl_user_list "
	sSql = sSql & " WHERE (distributionlistid IN (" & sdllist & ")) "
	sSql = sSql & " ORDER BY userid "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	'Loop thru email addresses sending email
	Do While Not oRs.EOF
		iRowCount = iRowCount + CLng(1)

		'This "If" is to handle skokie crashes. Change the user id found in the log to pick up from the crash point
		'If ( CLng(session("orgid")) = CLng(131) And CLng(oRs("userid")) > CLng(182789)) Or (CLng(session("orgid")) <> CLng(131)) Then 

		'SEND EMAIL		-- isValidEmail() is in common.asp
		If Not IsNull( oRs("useremail") ) Then 
			AddToLog "Sending # " & iRowCount
			AddToLog "From: "     & request("sFromName") & "[" & request("sFromEmail") & "]"
			AddToLog "Subject:"   & request("sSubjectLine")
			AddToLog "To: "       & LCase(oRs("useremail")) & " Userid: " & oRs("userid")

			If isValidEmail( LCase(oRs("useremail")) ) Then 
				subSendEmail request("sFromName"), request("sFromEmail"), request("sSubjectLine"), sHTMLBody, LCase(oRs("useremail")), oRs("userid"), clng(request("iEmailFormat")), request("containsHTML")
				AddToLog "Successful Send"
			Else 
				AddToLog "***** Email not sent due to invalid email format *****"
			End If 
		End If 

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' Sub subSendEmail( sFromName, sFromEmail, sSubjectLine, sHTMLBody, sToEmail, iUserId, iEmailFormat, iContainsHTML )
'-------------------------------------------------------------------------------------------------
Sub subSendEmail( sFromName, sFromEmail, sSubjectLine, sHTMLBody, sToEmail, iUserId, iEmailFormat, iContainsHTML )
	Dim sEmailFooter, lcl_email_from, lcl_email_to, lcl_email_subject, lcl_footer

	sEmailFooter = ""

	lcl_email_from     = sFromName & " <" & sFromEmail & ">"
	lcl_email_to       = sToEmail
	lcl_email_subject  = sSubjectLine

	If Not UserIsRootAdmin( session("userid") ) Then 
		lcl_footer = ""

		sEmailFooter = vbcrlf & vbcrlf & vbcrlf
		sEmailFooter = sEmailFooter & "<font style=""font-size:10pt;"">" & vbcrlf
		sEmailFooter = sEmailFooter & "<p>" & vbcrlf

		'Check to see if the org has overridden the default subscription footer
		If OrgHasDisplay( session("orgid"), "subscriptions_footer" ) Then 
			lcl_footer = GetOrgDisplay( session("orgid"), "subscriptions_footer" )
			lcl_footer = checkForCustomFields( lcl_footer, iUserID, sToEmail )
		End If 

		If Trim(lcl_footer) <> "" Then 
			sEmailFooter = sEmailFooter & lcl_footer
		Else 
			sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & " sent this e-mail to you because your Notification Preferences " & vbcrlf
			sEmailFooter = sEmailFooter & "indicate that you want to receive information from us. We will not request personal data (password, " & vbcrlf
			sEmailFooter = sEmailFooter & "credit card/bank numbers) in an e-mail. You are subscribed as " & sToEmail & ", " & vbcrlf
			sEmailFooter = sEmailFooter & "registered on " & session("egovclientwebsiteurl") & "." & vbcrlf
			sEmailFooter = sEmailFooter & "</p>" & vbcrlf
			sEmailFooter = sEmailFooter & "<p>" & vbcrlf
			sEmailFooter = sEmailFooter & "If you do not wish to receive further communications, sign into " & vbcrlf
			sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & " by clicking on the ""Login"" link found at the bottom of the " & vbcrlf
			sEmailFooter = sEmailFooter & session("egovclientwebsiteurl") & " home page and change your Notification Preferences or click " & vbcrlf
			sEmailFooter = sEmailFooter & "the link below to unsubscribe from this mailing list." & vbcrlf
			sEmailFooter = sEmailFooter & "</p>" & vbcrlf
			sEmailFooter = sEmailFooter & "<p>" & vbcrlf
			sEmailFooter = sEmailFooter & "<a href=""" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId & """>Click Here to Unsubscribe From our Mailing Lists</a>." & vbcrlf
		End If 

		sEmailFooter = sEmailFooter & "</p>" & vbcrlf
		sEmailFooter = sEmailFooter & "</font>" & vbcrlf
	End If 

	'Build the email body
	sHTMLBody = sHTMLBody & sEmailFooter

	If iContainsHTML = "" Then 
		iContainsHTML = "N"
	End If 

	If iEmailFormat < 3 Then 
		'include a plain text body
		If Not UserIsRootAdmin(session("userid")) Then 
			sHTMLBody = sHTMLBody & vbcrlf & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & iUserId 
		End If 
		'.TextBody = clearHTMLTags(sHTMLBody)  'This is in common.asp
	End If 

	lcl_email_htmlbody = BuildHTMLMessage( sHTMLBody, iContainsHTML )

	'Remove the name from the email address
	lcl_validate_email = formatSendToEmail( sFromEmail )

	'The function isValidEmail (found in common.asp) allows an email to simply have an "@" sign at the end of the email.
	'However, this will crash the application.  Check to see if the last character in the email entered is an "@".
	If lcl_validate_email <> "" And Right(lcl_validate_email,1) <> "@" Then 
		If isValidEmail( lcl_validate_email ) Then 
			'Send the email if it is valid.
			If iEmailFormat = 1 Then 
				sendEmail lcl_email_from, lcl_email_to, "", lcl_email_subject, "", lcl_email_htmlbody, "Y"
			Else 
				sendEmail lcl_email_from, lcl_email_to, "", lcl_email_subject, lcl_email_htmlbody, "", "Y"
			End If 
		End If 
	End If 

End Sub 


'-------------------------------------------------------------------------------------------------
' Function checkForCustomFields( p_value, p_userid, p_useremail )
'-------------------------------------------------------------------------------------------------
Function checkForCustomFields( p_value, p_userid, p_useremail )
	Dim lcl_return, lcl_unsubscribe_start, lcl_unsubscribe_end

	lcl_return = ""

	If p_value <> "" Then 
		lcl_return = p_value
		lcl_return = replace(lcl_return,"<<USER_EMAIL>>",p_useremail)
		lcl_return = replace(lcl_return,"<<ORGWEBSITE>>",session("egovclientwebsiteurl"))
		lcl_return = replace(lcl_return,"<<UNSUBSCRIBE>>","<a href=""" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & p_userid & """>Click Here to Unsubscribe From our Mailing Lists</a>")

		'Now check to see if they have a custom "Unsubscribe" text
		'The variable to be used in the "Edit Display" is: <<UNSUBSCRIBE_text goes here_UNSUBSCRIBE_END>>
		lcl_unsubscribe_start = "N"
		lcl_unsubscribe_end   = "N"

		'First check for the start of the "unsubscribe"
		If InStr( lcl_return, "<<UNSUBSCRIBE_" ) > 0 Then 
			lcl_unsubscribe_start = "Y"

			'If the "start" exists then check for the end of the "unsubscribe"
			If InStr( lcl_return, "_UNSUBSCRIBE_END>>" ) > 0 Then 
				lcl_unsubscribe_end = "Y"
			End If 
		End If 

		'If both the start and end of the "unsubscribe" exist then we can format them out
		'and build the unsubscribe link around the custom text.
		If lcl_unsubscribe_start = "Y" And lcl_unsubscribe_end = "Y" Then 
			lcl_return = Replace(lcl_return,"<<UNSUBSCRIBE_","<a href=""" & session("egovclientwebsiteurl") & "/subscriptions/subscribe_remove.asp?u=" & p_userid & """>")
			lcl_return = Replace(lcl_return,"_UNSUBSCRIBE_END>>","</a>")
		End If 
	End If 

	checkForCustomFields = lcl_return

End Function 


'-------------------------------------------------------------------------------------------------
' Sub updateSubscriptionInfoStatus( ByVal iDLLogID, ByVal sStatus )
'-------------------------------------------------------------------------------------------------
Sub updateSubscriptionInfoStatus( ByVal iDLLogID, ByVal sStatus )
	Dim sSql

	sSql = "UPDATE egov_class_distributionlist_log SET "
	sSql = sSql & " sendstatus = '" & UCase(dbsafe(sStatus)) & "', "
	sSql = sSql & " completedate = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
	sSql = sSql & " WHERE dl_logid = " & iDLLogID

	RunSQLStatement sSql 

End Sub 


'-------------------------------------------------------------------------------------------------
' Sub AddtoLog( sText )
'-------------------------------------------------------------------------------------------------
Sub AddtoLog( sText )
	Dim sSql

	' We do not want the local time here, as this is for us and we need to know when this was done in our time
	'sSql = "INSERT INTO subscriptionlog ( orgid, logentry ) VALUES ( " & session("orgid") & ", '" & replace(sText,"'","''") & "' )"
	sSql = "INSERT INTO subscriptionlog ( orgid, logentry ) VALUES ( " & session("orgid") & ", '" & dbsafe( sText ) & "' )"
	RunSQLStatement sSql 

End Sub 

'An inverse to Server.URLEncode -----------------------------------------------
function URLDecode(str)
	dim re
	set re = new RegExp

	str = Replace(str, "+", " ")
	
	re.Pattern = "%([0-9a-fA-F]{2})"
	re.Global = True
	URLDecode = re.Replace(str, GetRef("URLDecodeHex"))
end function

'Replacement function for the above -------------------------------------------
function URLDecodeHex(match, hex_digits, pos, source)
 	URLDecodeHex = chr("&H" & hex_digits)
end function


%>
