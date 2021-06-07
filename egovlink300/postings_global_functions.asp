<%
'------------------------------------------------------------------------------
function check_for_jobbid_categories(p_list_type)
 sSQL = "SELECT count(distributionlistid) AS total_count "
 sSQL = sSQL & " FROM egov_class_distributionlist "
 sSQL = sSQL & " WHERE orgid = '" & iorgid & "' "

 if p_list_type <> "" then
    sSQL = sSQL & " AND UPPER(distributionlisttype) = '" & UCASE(p_list_type) & "' "
 else
    sSQL = sSQL & " AND UPPER(distributionlisttype) NOT IN ('JOB','BID') "
 end if

 sSQL = sSQL & " AND distributionlistdisplay = 1 "

 set rs2 = Server.CreateObject("ADODB.Recordset")
 rs2.Open sSQL, Application("DSN"), 0, 1

 if not rs2.eof then
    if rs2("total_count") > 0 then
       lcl_return = "Y"
    else
       lcl_return = "N"
    end if
 else
    lcl_return = "N"
 end if

 check_for_jobbid_categories = lcl_return

end function

'------------------------------------------------------------------------------
function getCategoryName(p_dlistid,p_includeDesc)
  lcl_return = ""
  
  on error resume next

  sSQLc = "SELECT distributionlistname, distributionlistdescription "
  sSQLc = sSQLc & " FROM egov_class_distributionlist "
  sSQLc = sSQLc & " WHERE distributionlistid = '" & clng(p_dlistid) & "'"
  sSQLc = sSQLc & " AND orgid = " & iorgid

  set rsc = Server.CreateObject("ADODB.Recordset")
  rsc.Open sSQLc, Application("DSN"), 0, 1

  if not rsc.eof then
     lcl_return = rsc("distributionlistname")

    'Check to see if the description is to be included in the return value
     if UCASE(p_includeDesc) = "Y" then
        if rsc("distributionlistdescription") <> "" then
           lcl_return = lcl_return & "&nbsp;-&nbsp;" & rsc("distributionlistdescription")
        end if
     end if
  end if

  set rsc = nothing
  on error goto 0

  getCategoryName = lcl_return

end function

'-----------------------------------------------------------------------------
function getStatusName(p_status_id,p_list_type)
  lcl_return = ""

  if p_status_id <> "" AND p_list_type <> "" then
     sSQLs = "SELECT DISTINCT status_name "
     sSQLs = sSQLs & " FROM egov_statuses s "
     sSQLs = sSQLs & " WHERE status_id = " & CLng(p_status_id)
     sSQLs = sSQLs & " AND status_type = '" & p_list_type & "' "
     sSQLs = sSQLs & " AND active_flag = 'Y' "
     sSQLs = sSQLs & " AND orgid = " & iorgid

     set rss = Server.CreateObject("ADODB.Recordset")
     rss.Open sSQLs, Application("DSN"), 0, 1

     if not rss.eof then
        lcl_return = rss("status_name")
     end if
  end if

  set rss = nothing

  getStatusName = lcl_return

end function

'-----------------------------------------------------------------------------
function getPostingTitle(p_posting_id)
  sSQLp = "SELECT title "
  sSQLp = sSQLp & " FROM egov_jobs_bids "
  sSQLp = sSQLp & " WHERE posting_id = '" & CLng(p_posting_id) & "'"

  set rsp = Server.CreateObject("ADODB.Recordset")
  rsp.Open sSQLp, Application("DSN"), 0, 1

  if not rsp.eof then
     lcl_return = rsp("title")
  else
     lcl_return = ""
  end if

  set rsp = nothing

  getPostingTitle = lcl_return

end function

'------------------------------------------------------------------------------
function getStatusDefault(p_list_type)
  sSQLs1 = "SELECT status_id "
  sSQLs1 = sSQLs1 & " FROM egov_statuses "
  sSQLs1 = sSQLs1 & " WHERE orgid = " & iorgid
  sSQLs1 = sSQLs1 & " AND UPPER(status_type) = '" & UCASE(p_list_type) & "' "
  sSQLs1 = sSQLs1 & " AND default_status = 'Y' "

  set rss1 = Server.CreateObject("ADODB.Recordset")
  rss1.Open sSQLs1, Application("DSN"), 0, 1

  if not rss1.eof then
     lcl_return = rss1("status_id")
  else
     lcl_return = ""
  end if

  set rss1 = nothing

  getStatusDefault = lcl_return

end function

'------------------------------------------------------------------------------
function checkStatusType(p_list_type)

  lcl_exists = "N"

  if p_list_type <> "" then
     sSQL = "SELECT DISTINCT status_type "
     sSQL = sSQL & " FROM egov_statuses "
     sSQL = sSQL & " WHERE orgid = " & iorgid

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSQL, Application("DSN"), 0, 1

     lcl_exists = "N"

     if not rs.eof then
        while not rs.eof
           if UCASE(rs("status_type")) = UCASE(p_list_type) then
              lcl_exists = "Y"
           else
              lcl_exists = lcl_exists
           end if
           rs.movenext
        wend

        set rs = nothing
     else
        lcl_exists = lcl_exists
     end if
  end if

  checkStatusType = lcl_exists  

end function

'------------------------------------------------------------------------------
sub displaySignInLinks(p_list_label)

  response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td>" & vbcrlf
  response.write "          Must be signed in or register for an account before viewing "
  response.write            p_list_label & " Posting Specifications." & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "      <td align=""center"" width=""200"">" & vbcrlf
  response.write "          <a href=""user_login.asp"">Click here to Login</a><br />" & vbcrlf
  response.write "          or<br />" & vbcrlf
  response.write "          <a href=""register.asp?fromPostings=Y"">Click here to Register</a>." & vbcrlf
  response.write "      </td>" & vbcrlf
  response.write "  </tr>" & vbcrlf
  response.write "</table>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub displaySubscribeLink(p_dlistid)

  if p_dlistid <> "" AND dbready_number(p_dlistid) then
     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td>" & vbcrlf
     response.write "          You have not subscribed to the category (<strong>" & getCategoryName(p_dlistid,"") & "</strong>) "
     response.write "          associated to this posting." & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td align=""center"" width=""200"">" & vbcrlf
     response.write "          <input type=""button"" name=""subscribeButton"" id=""subscribeButton"" value=""Manage Subscriptions"" class=""button"" onclick=""location.href='manage_mail_lists.asp';"" />" & vbcrlf
     'response.write "          <a href=""manage_mail_lists.asp"">Click Here to Subscribe</a>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

  end if

end sub

'------------------------------------------------------------------------------
sub displayPostingsRequiredFields(iPostingID, iListType, iDLListID, iHasAllRequiredFields, iBusinessName, iWorkPhone)
  if iHasAllRequiredFields <> "Y" then
    'Check to see which fields have been populated.
     if iBusinessName <> "" then
        lcl_msg = ""
     else
        lcl_msg = "Business Name"
     end if

     if iWorkPhone <> "" then
       'NULL
     else
        if lcl_msg <> "" then
           lcl_msg = lcl_msg & " and Work Phone"
        else
           lcl_msg = "Work Phone"
        end if
     end if

     if lcl_msg <> "" then
        lcl_msg = "The following field(s) must be populated on your profile: <strong>" & lcl_msg & "</strong>"
     end if

     response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td>" & lcl_msg & "</td>" & vbcrlf
     response.write "      <td align=""center"" width=""200"">" & vbcrlf
     response.write "          <input type=""button"" name=""manageAccount"" id=""manageAccount"" value=""Manage Account"" class=""button"" onclick=""location.href='manage_account.asp?fromPostings=Y&listtype=" & iListType & "&posting_id=" & iPostingID & "&dllistid=" & iDLListID & "';"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "</table>" & vbcrlf

  end if

end sub

'------------------------------------------------------------------------------
function isCategoryAssigned(p_userid,p_categoryid)
  lcl_return     = False
  lcl_userid     = 0
  lcl_categoryid = 0

  if p_userid <> "" AND p_categoryid <> "" then
    'Validate input parameters
     if dbready_number(p_userid) then
        lcl_userid = p_userid
     end if

     if dbready_number(p_categoryid) then
        lcl_categoryid = p_categoryid
     end if

     sSQLd = "SELECT count(distributionlistid) as total_cnt "
     sSQLd = sSQLd & " FROM egov_class_distributionlist_to_user "
     sSQLd = sSQLd & " WHERE distributionlistid = '" & lcl_categoryid & "'"
     sSQLd = sSQLd & " AND userid = '" & lcl_userid & "'"

     set rsd = Server.CreateObject("ADODB.Recordset")
     rsd.Open sSQLd, Application("DSN"), 0, 1

     if rsd("total_cnt") > 0 then
        lcl_return = True
     end if

     set rsd = nothing

  end if

  isCategoryAssigned = lcl_return

end function

'-----------------------------------------------------------------------------
sub getPostingInfo(ByVal p_posting_id, ByRef lcl_jobbid_id, ByRef lcl_title, ByRef lcl_description, ByRef lcl_enddate, ByRef lcl_statusname)
  lcl_jobbid_id   = ""
  lcl_title       = ""
  lcl_description = ""
  lcl_enddate     = ""
  lcl_statusname  = ""

  if p_posting_id <> "" then
     sSQLp = "SELECT jb.jobbid_id, jb.title, jb.description, jb.end_date, "
     sSQLp = sSQLp & " (select status_name from egov_statuses s where s.status_id = jb.status_id) AS statusname "
     sSQLp = sSQLp & " FROM egov_jobs_bids jb "
     sSQLp = sSQLp & " WHERE jb.posting_id = " & p_posting_id

     set rsp = Server.CreateObject("ADODB.Recordset")
     rsp.Open sSQLp, Application("DSN"), 0, 1

     if not rsp.eof then
        lcl_jobbid_id   = rsp("jobbid_id")
        lcl_title       = rsp("title")
        lcl_description = rsp("description")
        lcl_statusname  = rsp("statusname")

       'Format end date
        if rsp("end_date") = "1/1/1900" then
           lcl_enddate = ""
        else
           lcl_enddate = rsp("end_date")
        end if

     end if

     set rsp = nothing
  end if

end sub

'------------------------------------------------------------------------------
function getUserEmail(p_userid)
  lcl_return = ""

  if p_userid <> "" then
     sSQL = "SELECT useremail "
     sSQL = sSQL & " FROM egov_users "
     sSQL = sSQL & " WHERE orgid = " & iorgid
     sSQL = sSQL & " AND userid = " & p_userid

     set oUserEmail = Server.CreateObject("ADODB.Recordset")
     oUserEmail.Open sSQL, Application("DSN"), 3, 1

     if not oUserEmail.eof then
        lcl_return = oUserEmail("useremail")
     end if

  end if

  set oUserEmail = nothing

  getUserEmail = lcl_return

end function

'------------------------------------------------------------------------------
function getPostingsAdminEmail()
  lcl_return = ""

  sSQL = "SELECT postings_userbids_notifyemail "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE orgid = " & iorgid

  set oAdminEmail = Server.CreateObject("ADODB.Recordset")
  oAdminEmail.Open sSQL, Application("DSN"), 3, 1

  if not oAdminEmail.eof then
     lcl_return = oAdminEmail("postings_userbids_notifyemail")
  end if

  set oAdminEmail = nothing

  getPostingsAdminEmail = lcl_return

end function

'------------------------------------------------------------------------------
sub sendUploadEmail(p_posting_id, p_useremail, p_adminemail, p_uploadsDirVar, p_upload_filename, p_uploadid, p_userLabel)
   lcl_sendtoemail = ""

   if p_useremail <> "" then
      lcl_sendtoemail = p_useremail
   elseif p_adminemail <> "" then
      lcl_sendtoemail = p_adminemail
   end if

   if lcl_sendtoemail <> "" AND p_posting_id <> "" then
      lcl_orgname = UCASE(oOrg.GetOrgName())

      datGMTDateTime = DateAdd("h",5,Now())
      datOrgDateTime = DateAdd("h",iTimeOffset,datGMTDateTime)
      datCurrentDate = Now()

     'Get the posting info
      getPostingInfo p_posting_id, lcl_jobbid_id, lcl_title, lcl_description, lcl_enddate, lcl_statusname

     'Build email message body
      sMsg = "<p>This automated message was sent by the " & lcl_orgname & ".  Do not reply to this message.</p>" & vbcrlf 

      if p_useremail <> "" then
         sMsg = sMsg & "<p>Thank you for submitting your information to " & lcl_orgname & " on " & datOrgDateTime & ".</p>" & vbcrlf
      else
         sMsg = sMsg & "<p>A user has uploaded a bid for the following bid posting.</p>" & vbcrlf
      end if

			   sMsg = sMsg & "<p><strong>BID NUMBER: </strong>"  & lcl_jobbid_id   & "</p>" & vbcrlf
			   sMsg = sMsg & "<p><strong>TITLE: </strong>"       & lcl_title       & "</p>" & vbcrlf
			   sMsg = sMsg & "<p><strong>DESCRIPTION: </strong>" & lcl_description & "</p>" & vbcrlf
			   sMsg = sMsg & "<p><strong>LABEL: </strong>"       & p_userLabel     & "</p>" & vbcrlf
			   sMsg = sMsg & "<p><strong>UPLOAD ID: </strong>"   & p_uploadid      & "</p>" & vbcrlf & vbcrlf


		
  			'Send email
      if isValidEmail(lcl_sendtoemail) then
		sendEmail lcl_orgname & " Bid Upload Notification <webmaster@eclink.com>", lcl_sendtoemail, "", lcl_orgname & " BID Upload Notification", sMsg, clearHTMLTags(sMsg), "N"
      else
         ErrorCode = 1
      end if
			
  			'Add to email queue if unsuccessful
		  		if ErrorCode <> 0 then
    					sMsg = Left(sMsg,5000)
		    			SendToAdd = lcl_sendtoemail
				    	fnPlaceEmailinQueue Application("SMTP_Server"),lcl_orgname & " Bid Uploads <webmaster@eclink.com>","webmaster@eclink.com",lcl_sendtoemail,lcl_orgname & " - UPLOAD BID EMAIL",1,sMsg,1,-1
  				end if

   			if ErrorCode <> 0 then
			   	 'Add logging code
     				response.write "The request has been logged but there was an error sending an email notice to you.  You will not receive an email notice.<br /><br /><br />"
     				bMailSent1 = False
   			end if
  end if

end sub

'----------------------------------------------------------------------------- 
Function fnPlaceEmailinQueue(sHost,sFromName,sFromEmail,sSendEmail,sSubject,iBodyFormat,sBodyMessage,iPriority,iErrorCode)
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

sub dtb_debug(p_value)
  if p_value <> "" then
     sSQLi = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"
     set rsi = Server.CreateObject("ADODB.Recordset")
     rsi.Open sSQLi, Application("DSN"), 3, 1

     set rsi = nothing
  end if
end sub
%>
