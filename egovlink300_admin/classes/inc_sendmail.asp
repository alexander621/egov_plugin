<%
'------------------------------------------------------------------------------
sub subProcessEmail(iDL_logID)
	dim iRowCount

	iRowCount = CLng(0) 

'Get the distribution list(s)
	sdllist = replace(request("Maillist"),"X",",") ' BUILD COMMA SEPARATE LIST
	iLength = LEN(sdllist)
	sdllist = LEFT(sdllist,(iLength - 1)) ' TRIM TRAILING COMMA

	'GET LIST OF UNIQUE EMAIL ADDRESSES
	sSQL = "SELECT MAX(userid) as userid, useremail "
	sSQL = sSQL & " FROM egov_dl_user_list "
	sSQL = sSQL & " WHERE (distributionlistid IN (" & sdllist & ")) "
 sSQL = sSQL & " GROUP BY useremail ORDER BY MAX(userid) "

 set oEmail = Server.CreateObject("ADODB.Recordset")
 oEmail.Open sSQL, Application("DSN"), 3, 1

 TotalEmails = oEmail.RecordCount
 
 if showGraph then
 %>
 <table cellpadding=0 cellspacing=0 border=2 width="500" id="progresstable">
 	<tr>
		<td style="height:50px;">
			<div style="position:relative;margin-top:-25px;">
				<div id="bar" style="background:red;float:left;z-index:2;width:0%;height:50px;position:absolute;"></div>
				<div id="percent" style="float:left;width:100%;text-align:center;z-index:3;position:absolute;margin-top:16px;">0%</div>
			</div>
		</td>
	</tr>
	<tr>
		<td align="center">Sent <span id="currentcount">0</span> of <%=TotalEmails%></td>
	</tr>
 </table>
 <script>
	function u(percent, currentcount)
	{
		document.getElementById("bar").style.width = percent;
		document.getElementById("percent").innerHTML = percent;
		document.getElementById("currentcount").innerHTML = currentcount;
	}
 </script>
 <%
 response.flush
 end if

'Loop thru email addresses sending email
	currentcount = 0
	do while not oEmail.eof
  		iRowCount = iRowCount + CLng(1)

		'SEND EMAIL		-- isValidEmail() is in common.asp
 		if not IsNull( oEmail("useremail") ) then
	   		AddToLog "(dl_logid: " & iDL_logID & ") Sending # " & iRowCount
   			AddToLog "(dl_logid: " & iDL_logID & ") From: "     & request("sFromName") & "[" & request("sFromEmail") & "]"
		   	AddToLog "(dl_logid: " & iDL_logID & ") Subject:"   & request("sSubjectLine")
   			AddToLog "(dl_logid: " & iDL_logID & ") To: "       & LCase(oEmail("useremail")) & " Userid: " & oEmail("userid")
               		AddtoLog "(dl_logid: " & iDL_logID & ") Subscription Log ID: " & iDL_logID

  	'if intOrgID = "37" then
		'response.write "&&&" & request("sFromName") & "###"
		'response.end
	'end if

   			if isValidEmail( LCase(oEmail("useremail")) ) then
  	   			subSendEmail request("sFromName"), request("sFromEmail"), request("sSubjectLine"), request("sHTMLBody"), LCase(oEmail("useremail")), oEmail("userid"), clng(request("iEmailFormat")), request("containsHTML"), sdllist
      				AddToLog "(" & iDL_logID & ") Successful Send"
       			else
  	   			AddToLog "(" & iDL_logID & ") ***** Email not sent due to invalid email format *****"
       			end if

 		end if

		
		if showGraph then 
    			currentcount = currentcount + 1
			Percent = FormatNumber((currentcount/TotalEmails)*100,0)
			response.write "<script>u('" & Percent & "%','" & currentcount & "')</script>" & vbcrlf
			response.flush
		end if

   		oEmail.movenext
	loop

	if showGraph then 
		response.write "<script>document.getElementById('progresstable').style.display='none';</script>"
		response.flush
		'response.end
	end if

	oEmail.Close
	set oEmail = nothing 

end sub


'------------------------------------------------------------------------------
sub subSendEmail( sFromName, sFromEmail, sSubjectLine, sHTMLBody, sToEmail, iUserId, iEmailFormat, iContainsHTML, iDLListID )
 	dim sEmailFooter

	 sEmailFooter = ""

  lcl_email_from     = sFromName & " <" & sFromEmail & ">"
  lcl_email_to       = sToEmail
  lcl_email_subject  = sSubjectLine


 	if not userisrootadmin(intUserID) or intOrgID <> "113" then
     lcl_footer = ""

 		  sEmailFooter = vbcrlf & vbcrlf & vbcrlf
     sEmailFooter = sEmailFooter & "<font style=""font-size:10pt;"">" & vbcrlf
     sEmailFooter = sEmailFooter & "<p>" & vbcrlf

    'Check to see if the org has overridden the default subscription footer
     if OrgHasDisplay(intOrgID,"subscriptions_footer") then
        lcl_footer = GetOrgDisplay(intOrgID,"subscriptions_footer")
        lcl_footer = checkForCustomFields(lcl_footer,iUserID,sToEmail)
     end if

     if trim(lcl_footer) <> "" then
        sEmailFooter = sEmailFooter & lcl_footer
     else
        'Get the name of the list(s) to be unsubscribed from
         lcl_display_dlids   = ""
         lcl_display_dlnames = ""

         'sSQL = "SELECT distributionlistname "
         'sSQL = sSQL & " FROM egov_class_distributionlist "
         'sSQL = sSQL & " WHERE distributionlistid IN (" & iDLListID& ") "
         'sSQL = sSQL & " ORDER BY distributionlistname "

         sSQL = "SELECT dl.distributionlistid, "
         sSQL = sSQL & " dl.distributionlistname "
         sSQL = sSQL & " FROM egov_class_distributionlist dl, "
         sSQL = sSQL &      " egov_class_distributionlist_to_user dltu "
         sSQL = sSQL & " WHERE dl.distributionlistid = dltu.distributionlistid "
         sSQL = sSQL & " AND dl.distributionlistid IN (" & iDLListID& ") "
         sSQL = sSQL & " AND dltu.userid = " & iUserID
         sSQL = sSQL & " ORDER BY dl.distributionlistname "

         set oGetDLNames = Server.CreateObject("ADODB.Recordset")
         oGetDLNames.Open sSQL, Application("DSN"), 0, 1

         if not oGetDLNames.eof then
            do while not oGetDLNames.eof

              'Build the "display" distribution list IDs (for the unsubscribe URL)
               if lcl_display_dlids = "" then
                  lcl_display_dlids = oGetDLNames("distributionlistid")
               else
                  lcl_display_dlids = lcl_display_dlids & "," & oGetDLNames("distributionlistid")
               end if

              'Build the "display" distribution list names
               if lcl_display_dlnames = "" then
                  lcl_display_dlnames = oGetDLNames("distributionlistname")
               else
                  lcl_display_dlnames = lcl_display_dlnames & ", " & oGetDLNames("distributionlistname")
               end if

               oGetDLNames.movenext
            loop
         end if

         oGetDLNames.close
         set oGetDLNames = nothing

	'Get WP Live Status & Public URL
	blnWPLive = false
	strPublicURL = ""
	sSQL = "SELECT wpLive,OrgPublicWebsiteURL FROM organizations WHERE orgid = '" & intOrgID & "'"
	set oO = Server.CreateObject("ADODB.RecordSet")
	oO.Open sSQl, Application("DSN"), 3, 1
	if not oO.EOF then
		blnWPLive = oO("wpLive")
		strPublicURL = oO("OrgPublicWebsiteURL")
	end if
	oO.Close
	Set oO = Nothing

	strConfirmationID = ""
	sSQL = "SELECT subscription_confirmid FROM egov_users WHERE userid = '" & iUserID & "'"
	set oU = Server.CreateObject("ADODB.RecordSet")
	oU.Open sSQl, Application("DSN"), 3, 1
	if not oU.EOF then
		strConfirmationID = oU("subscription_confirmid")

		if isnull(strConfirmationID) or strConfirmationID = "" then
			'encode the userid for compare
			strCode = createHashedPassword("thisiscomplex" & iUserID)
			strConfirmationID = strCode & "|" & iUserID
		end if
	end if
	oU.Close
	Set oU = Nothing

	subURL = getOrganization_WP_URL(intOrgID, "wp_subscriptions_url")
	if subURL = "" then subURL = strEGovURL & "/manage_mail_lists.asp"

        sEmailFooter = sEmailFooter & strOrgName & " sent this e-mail to you because your Notification Preferences "
        sEmailFooter = sEmailFooter & "indicate that you want to receive information from us. We will not request personal data (password, "
        sEmailFooter = sEmailFooter & "credit card/bank numbers) in an e-mail. You are subscribed as " & sToEmail & ", "
        sEmailFooter = sEmailFooter & "registered on " & strOrgName & " (<a href=""" & subURL & """>" & subURL & "</a>)."
        sEmailFooter = sEmailFooter & "</p>"
        sEmailFooter = sEmailFooter & "<p>"
        sEmailFooter = sEmailFooter & "<strong>Click Here to Unsubscribe From this List(s):</strong>" & vbcrlf
        sEmailFooter = sEmailFooter & "You will be removed from the following lists: " & lcl_display_dlnames & "<br />"
	if strConfirmationID <> "" and not isnull(strConfirmationID) and blnWPLive then
        	sEmailFooter = sEmailFooter & "<a href=""" & subURL & "#unsubscribe/" & strConfirmationID & """>Unsubscribe</a>.<br />" & vbcrlf
	else
        	sEmailFooter = sEmailFooter & "<a href=""" & strEGovURL & "/subscriptions/subscribe_remove.asp?u=" & iUserId & "&dl=" & iDLListID & """>Unsubscribe</a>.<br />" & vbcrlf
	end if
        sEmailFooter = sEmailFooter & "</p>"
        sEmailFooter = sEmailFooter & "<p>"
        sEmailFooter = sEmailFooter & "<strong>Manage Subscriptions:</strong>" & vbcrlf
        sEmailFooter = sEmailFooter & "If you do not wish to receive further communications, or you wish to view "
        sEmailFooter = sEmailFooter & "and/or modify which lists you are subscribed to, simply click the link below.<br />"
        sEmailFooter = sEmailFooter & "<a href=""" & subURL & """>" & subURL & "</a>"
     end if

     sEmailFooter = sEmailFooter & "</p>"
     sEmailFooter = sEmailFooter & "</font>"
 	end if

 'Build the email body
 	sHTMLBody = sHTMLBody & sEmailFooter

  'if iContainsHTML = "" then
  '   iContainsHTML = "N"
  'end if
  iContainsHTML = "Y"


  lcl_email_htmlbody = BuildHTMLMessage(sHTMLBody, iContainsHTML)

 'Remove the name from the email address
  lcl_validate_email = formatSendToEmail(sFromEmail)

 'The function isValidEmail (found in common.asp) allows an email to simply have an "@" sign at the end of the email.
 'However, this will crash the application.  Check to see if the last character in the email entered is an "@".
  'dblogging "HERE 1: " & lcl_validate_email
  if lcl_validate_email <> "" AND RIGHT(lcl_validate_email,1) <> "@" then
  	'dblogging "HERE 2: " & isValidEmail(lcl_validate_email)
     if isValidEmail(lcl_validate_email) then

       'Send the email if it is valid.

        if iEmailFormat = 1 then
           sendEmail lcl_email_from, lcl_email_to, "", lcl_email_subject, "", lcl_email_htmlbody, "Y"
        else
           sendEmail lcl_email_from, lcl_email_to, "", lcl_email_subject, lcl_email_htmlbody, "", "Y"
        end if
     end if
  end if

end sub

Sub dblogging( sText )
	sSQL = "INSERT INTO zdblogging (dbmessage) VALUES('" & sText & "')"
	RunSQLStatement( sSQL )

End Sub

'----------------------------------------------------------------------------------------
Sub AddtoLog( sText )
    ' WRITES SUPPLIED TEXT TO FILE WITH DATETIME
'	Set oFSO = Server.Createobject("Scripting.FileSystemObject")
'    Set oFile = oFSO.GetFile(Application("SubscriptionLog"))
'    Set oText = oFile.OpenAsTextStream(8)
'    oText.WriteLine (Now() & Chr(9) & sText)
'    oText.Close
    
'    Set oText = Nothing
'    Set oFile = Nothing
 '   Set oFSO = Nothing

	Dim sSql

	sSql = "INSERT INTO subscriptionlog ( orgid, logentry ) VALUES ( " & intOrgID & ", '" & replace(sText,"'","''") & "' )"
	RunSQLStatement( sSql )

End Sub 
'------------------------------------------------------------------------------
sub updateJustSubscriptionInfoStatus(iDLLogID, iStatus)

  sSQL = "UPDATE egov_class_distributionlist_log SET "
  sSQL = sSQL & " sendstatus = '"   & UCASE(dbsafe(iStatus)) & "' "
  sSQL = sSQL & " WHERE dl_logid = " & iDLLogID

	 set oDLLogUpdate = Server.CreateObject("ADODB.Recordset")
		oDLLogUpdate.Open sSQL, Application("DSN"), 0, 1

  set oDLLogUpdate = nothing

end sub
sub updateSubscriptionInfoStatus(iDLLogID, iStatus)

  sSQL = "UPDATE egov_class_distributionlist_log SET "
  sSQL = sSQL & " sendstatus = '"   & UCASE(dbsafe(iStatus)) & "', "
  sSQL = sSQL & " completedate = '" & now                    & "' "
  sSQL = sSQL & " WHERE dl_logid = " & iDLLogID

	 set oDLLogUpdate = Server.CreateObject("ADODB.Recordset")
		oDLLogUpdate.Open sSQL, Application("DSN"), 0, 1

  set oDLLogUpdate = nothing

end sub

'------------------------------------------------------------------------------
sub clearSubscriptionInfoSchedule(iDLLogID)

  sSQL = "UPDATE egov_class_distributionlist_log SET "
  sSQL = sSQL & " scheduledDateTime = NULL "
  sSQL = sSQL & " WHERE dl_logid = " & iDLLogID

	 set oDLLogUpdate = Server.CreateObject("ADODB.Recordset")
		oDLLogUpdate.Open sSQL, Application("DSN"), 0, 1

  set oDLLogUpdate = nothing

end sub
sub SendPushNotification()
    'Send Push Notification
    sSQL = "SELECT o.orgid,'j' + CONVERT(varchar(10),o.orgid) + 'c' + CONVERT(varchar(10),distributionlistid) as channel " _
		& " FROM egov_class_distributionlist dl " _
		& " INNER JOIN egovlinkRegistry.dbo.Orgs o ON dl.orgid = o.legacyorgid " _
		& " WHERE dl.orgid = '" & intOrgID & "' and dl.distributionlistid IN (" & request.form("sendlist") & ") "
    Set oPN = Server.CreateObject("ADODB.RecordSet")
    oPN.Open sSQL, Application("DSN"), 3, 1
    if not oPN.EOF then
	    'channel = oPN("channel")
	    jurisdiction_id = oPN("orgid")

	    channels = ""
	    Do While Not oPN.EOF
	    	channels = channels & """" & oPN("channel") & ""","
	    	oPN.MoveNext
	    loop
	    channels = left(channels,len(channels)-1)
    	    postData = "{""jurisdiction_id"":""" & jurisdiction_id & """,""channels"":[" & channels & "],""subject"":""" & replace(request.form("sSubjectLine"),"""","\""") & """,""body"":""" & replace(request.form("sHTMLBody"),"""","\""") & """,""key"":""value"" }"
	    'response.write postData
	    'response.end

	    Set xmlHttp = Server.CreateObject("Microsoft.XMLHTTP") 
	    xmlHttp.Open "POST", "http://registry2.eclinkhost.com/messages/send", False
	    xmlHttp.setRequestHeader "Content-Type", "application/json"
	    xmlHttp.Send postData
	    set xmlHttp = Nothing 
    end if
    oPN.Close
    Set oPN = Nothing
end sub

'------------------------------------------------------------------------------
function checkForCustomFields(p_value,p_userid,p_useremail)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,"<<USER_EMAIL>>",p_useremail)
     lcl_return = replace(lcl_return,"<<ORGWEBSITE>>",strEGovURL)
     lcl_return = replace(lcl_return,"<<UNSUBSCRIBE>>","<a href=""" & strEGovURL & "/subscriptions/subscribe_remove.asp?u=" & p_userid & """>Click Here to Unsubscribe From our Mailing Lists</a>")

    'Now check to see if they have a custom "Unsubscribe" text
    'The variable to be used in the "Edit Display" is: <<UNSUBSCRIBE_text goes here_UNSUBSCRIBE_END>>
     lcl_unsubscribe_start = "N"
     lcl_unsubscribe_end   = "N"

    'First check for the start of the "unsubscribe"
     if instr(lcl_return,"<<UNSUBSCRIBE_") > 0 then
        lcl_unsubscribe_start = "Y"

       'If the "start" exists then check for the end of the "unsubscribe"
        if instr(lcl_return,"_UNSUBSCRIBE_END>>") > 0 then
           lcl_unsubscribe_end = "Y"
        end if
     end if

    'If both the start and end of the "unsubscribe" exist then we can format them out
    'and build the unsubscribe link around the custom text.
     if lcl_unsubscribe_start = "Y" AND lcl_unsubscribe_end = "Y" then
        lcl_return = replace(lcl_return,"<<UNSUBSCRIBE_","<a href=""" & strEGovURL & "/subscriptions/subscribe_remove.asp?u=" & p_userid & """>")
        lcl_return = replace(lcl_return,"_UNSUBSCRIBE_END>>","</a>")
     end if
  end if

  checkForCustomFields = lcl_return

end function



%>
<!-- #include file="../../egovlink300_global/includes/inc_passencryption.asp" //-->
