<!-- #include file="../../egovlink300_global/includes/inc_passencryption.asp" //-->
<%
'------------------------------------------------------------------------------
function setupScreenMsg(iSuccess)

  lcl_return = ""

  if iSuccess <> "" then
     iSuccess = UCASE(iSuccess)

     if iSuccess = "SU" then
        lcl_return = "Your subscription selection(s) have been saved successfully."
     elseif iSuccess = "SA" then
        lcl_return = "Your subscription selection(s) have been submitted successfully."
     elseif iSuccess = "EXISTS_BAD_PWD" then
        lcl_return = "An account already exists with this email, but the password does not match."
     elseif iSuccess = "REQUIRED_PWD" then
        lcl_return = "A password is required."
     elseif iSuccess = "BAD_PWD" then
        lcl_return = "The password entered is incorrect."
     elseif iSuccess = "NOT_EXISTS" then
        lcl_return = "We do not have an account with this email address."
     elseif iSuccess = "UNSUBSCRIBED" then
        lcl_return = "ALL subscriptions have been removed."
     elseif iSuccess = "AJAX_ERROR" then
        lcl_return = "ERROR: An error has during the AJAX routine..."
     end if
  end if

  setupScreenMsg = lcl_return

end function

'------------------------------------------------------------------------------
sub checkUserExists( ByVal p_userid, ByVal p_useremail, ByVal p_orgid, ByVal p_confirmid, ByRef lcl_user_id, ByRef lcl_user_email, ByRef lcl_user_password, ByRef lcl_user_address )
	Dim sSQL, okToRun, iUserEmail, iConfirmID, lcl_userid

	lcl_userid        = ""
	lcl_user_email    = ""
	lcl_user_password = ""
	lcl_user_address  = ""
	okToRun = False 

	iConfirmID = ""
	iUserEmail = ""

	if p_confirmid <> "" then
		iConfirmID = p_confirmid
		iConfirmID = dbsafe(iConfirmID)
		iConfirmID = "'" & iConfirmID & "'"
	end if

	' a very popular intrusion attack uses single quotes as the email address, so block these here
	p_useremail = Replace(p_useremail,"'","")
	if p_useremail <> "" then
		iUserEmail = p_useremail
		iUserEmail = ucase(iUserEmail)
		iUserEmail = dbsafe(iUserEmail)
		iUserEmail = "'" & iUserEmail & "'"
	end if

	'Check to see if user exists
	sSQL = "SELECT userid, useremail, password, useraddress FROM egov_users "

	'If a subscription confirmation id is passed in then try and find the user on subscription_confirmid
	'If "no" then if a userid is passed in search on userid
	'If "no" then search on email address.
	if iConfirmID <> "" then
		sSQL = sSQL & " WHERE subscription_confirmid = " & iConfirmID
		okToRun = True 
	else
		if p_userid <> "" then
			sSQL = sSQL & " WHERE userid = " & CLng(p_userid)
			okToRun = True 
		Else
			If iUserEmail <> "" And InStr( iUserEmail, "@" ) > 0 Then 
				sSQL = sSQL & " WHERE UPPER(useremail) = " & iUserEmail
				sSQL = sSQL & " AND orgid = " & CLng(p_orgid)
				okToRun = True 
			End If 
		end if
	end If

	session("checkUserExists_SQL") = sSQL

	If okToRun Then 
		set oUser = Server.CreateObject("ADODB.Recordset")
		oUser.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

		if not oUser.eof then
			lcl_user_id       = oUser("userid")
			lcl_user_email    = oUser("useremail")
			lcl_user_password = oUser("password")
			lcl_user_address  = oUser("useraddress")
		end if

		oUser.close
		set oUser = Nothing
	End If 

end sub

'------------------------------------------------------------------------------
sub SendSpamFlag(sFromEmail, sPassword, sTextinput, sOrgID, sIPAddress)

 'Setup the email variables
  sOrgName    = getOrgName(sOrgID)
  lcl_from    = sOrgName & " E-GOV WEBSITE <webmaster@eclink.com>"
  lcl_to      = "egovsupport@eclink.com"
  lcl_subject = sOrgName & " - E-Gov (Possible Subscription Spam Submission)"

 'Build the message
 	sMsgBody = "Possible SUBSCRIPTIONS spam submitted to " & sOrgName & ".<br />" & vbcrlf
 	sMsgBody = sMsgBody & "<p>Subscriber Email: "   & sFromEmail  & "<br />" & vbcrlf
 	sMsgBody = sMsgBody & "Subscriber Password: "   & sPassword   & "<br />" & vbcrlf
 	sMsgBody = sMsgBody & "Hidden field contains: " & sTextinput  & "<br />" & vbcrlf
  sMsgBody = sMsgBody & "IP Address: "            & sIPAddress  & "</p>" & vbcrlf

 'Send the email
  sendEmail lcl_from, lcl_to, "", lcl_subject, sMsgBody, "", "Y"

end sub

'------------------------------------------------------------------------------
sub InsertMailList( ByVal iUserID, ByVal iListID )

	dim lcl_uid, lcl_lid, sSQL, oInsertList

	lcl_uid = 0
	lcl_lid = 0

	if iUserID <> "" then
		lcl_uid = CLng(iUserID)
	end if

	if iListID <> "" then
		lcl_lid = CLng(iListID)
	end if

	sSQL = "INSERT INTO egov_class_distributionlist_to_user ("
	sSQL = sSQL & "userid, "
	sSQL = sSQL & "distributionlistid"
	sSQL = sSQL & ") VALUES ("
	sSQL = sSQL & lcl_uid & ", "
	sSQL = sSQL & lcl_lid
	sSQL = sSQL & ")"

	set oInsertList = Server.CreateObject("ADODB.Recordset")
	oInsertList.Open sSQL, Application("DSN"), 0, 1

	set oInsertList = nothing

end sub

'------------------------------------------------------------------------------
sub SendConfirmation( sToAddress, iUserid, sPassword, iConfirmID )

'Set up mail variables
 lcl_from    = sOrgName & " E-Gov Website <noreplies@egovlink.com>"
 lcl_to      = sToAddress
 lcl_subject = sOrgName & " Subscriptions Signup"
	sTextBody = ""
	sHTMLBody = ""

'Build the HTML message
 'sHTMLBody = "<p>(<a href=""" & Application("SUBSCRIBE_URL") & "unsubscribe.asp?c=" & iConfirmID & """>Unsubscribe at " & Application("SUBSCRIBE_URL") & "unsubscribe.asp?c=" & iConfirmID & "</a>)</p>" & vbcrlf
 sHTMLBody = "<p>(<a href=""" & sEgovWebsiteURL & "/unsubscribe.asp?c=" & iConfirmID & """>Unsubscribe at " & sEgovWebsiteURL & "/unsubscribe.asp?c=" & iConfirmID & "</a>)</p>" & vbcrlf
 sHTMLBody = sHTMLBody & "<p>Thank you for signing up for our email subscriptions.</p>" & vbcrlf
 sHTMLBody = sHTMLBody & "<p>You have signed up for the following subscriptions.</p>" & vbcrlf & vbcrlf
 sHTMLBody = sHTMLBody & "<p>"

	for each list in request("maillist")
    	sHTMLBody = sHTMLBody & vbcrlf & GetDistributionListName( list )
 next

	sHTMLBody = sHTMLBody & "</p>" & vbcrlf & vbcrlf
    sHTMLBody = sHTMLBody & "<p>Maintain your subscriptions using the following URL:</p>" & vbcrlf & vbcrlf
 	'sHTMLBody = sHTMLBody & "<p><a href=""" & sEgovWebsiteURL & "/subscriptions/subscribe.asp?u=" & iUserid & """>" & sEgovWebsiteURL & "/subscriptions/subscribe.asp?u=" & iUserid & "</a></p>" & vbcrlf & vbcrlf
	sHTMLBody = sHTMLBody & "<p><a href=""" & sEgovWebsiteURL & "/subscriptions/subscribe.asp"">" & sEgovWebsiteURL & "/subscriptions/subscribe.asp</a></p>" & vbcrlf & vbcrlf

'Format the Text Message
	sTextBody = sHTMLBody
	sTextBody = clearHTMLTags( sTextBody )

'Send the email
 sendEmail lcl_from, lcl_to, "", lcl_subject, sHTMLBody , sTextBody, "Y"

end sub

'------------------------------------------------------------------------------
sub ClearMailList(p_userid)

  if p_userid <> "" then
     sSQL = "DELETE FROM egov_class_distributionlist_to_user WHERE userid = " & p_userid

     set oDeleteList = Server.CreateObject("ADODB.Recordset")
     oDeleteList.Open sSQL, Application("DSN"), 0, 1

    set oDeleteList = nothing
  end if

end sub

'------------------------------------------------------------------------------
function ProcessRecords( ByVal p_orgid, ByVal p_useremail, ByVal p_userpassword, ByVal p_residenttype, ByVal p_userregistered, ByVal p_headofhousehold )
	 Dim iRelationshipId, sSql

	'Set database field lengths
	lcl_length_useremail    = 512
	lcl_length_userpassword = 50
	lcl_length_residenttype = 1
	lcl_length_firstname    = 50
	lcl_length_lastname     = 50
	lcl_length_relationship = 20

	lcl_useremail       = dbready_string( p_useremail, lcl_length_useremail )
	'lcl_userpassword    = dbready_string( p_userpassword, lcl_length_userpassword )
	lcl_userpassword    = createHashedPassword( p_userpassword )
	lcl_residenttype    = dbready_string( p_residenttype, lcl_length_residenttype )
	lcl_userregistered  = p_userregistered
	lcl_headofhousehold = p_headofhousehold
	iRelationshipId     = GetDefaultRelationShipId( p_orgid )

	'If the user exists then attempt to bring in their email/password
	checkUserExists "", useremail, p_orgid, "", lcl_user_id, lcl_user_email, lcl_user_password, lcl_user_address

	if lcl_user_id <> "" then

		'Build query to UPDATE egov_users
		sSql = "UPDATE egov_users SET "
		sSql = sSql & " orgid = " & p_orgid & ", "
		sSql = sSql & " residenttype = '" & lcl_residenttype & "', "
		sSql = sSql & " headofhousehold = "  & lcl_headofhousehold & ", "
		sSql = sSql & " useremail = '" & lcl_useremail & "', "
		sSql = sSql & " password = '" & lcl_userpassword & "' "
		sSql = sSql & " WHERE userid = " & lcl_user_id

		RunSQLStatement sSql  'In Common.asp

	Else 
		' try to filter out some spam with this check
		If InStr( lcl_useremail, "@" ) > 0 Then
			'Build query to INSERT record into egov_users
			sSql = "INSERT INTO egov_users ( orgid, useremail, password, userregistered, residenttype, headofhousehold, relationshipid"
			sSql = sSql & ") VALUES ("
			sSql = sSql & iorgid & ", "
			sSql = sSql & "'" & lcl_useremail & "', "
			sSql = sSql & "'" & lcl_userpassword & "', "
			sSql = sSql & lcl_userregistered & ", "
			sSql = sSql & "'" & lcl_residenttype & "', "
			sSql = sSql & lcl_headofhousehold & ", "
			sSql = sSql & iRelationshipId
			sSql = sSql & ")"
			'	response.write sSql & "<br />"
			'	response.End 
		 
			lcl_userid = RunIdentityInsertStatement( sSql )		' In Common.asp
			' response.write "lcl_userid = " & lcl_userid & "<br />"
		
'			set oUser = Server.CreateObject("ADODB.Recordset")
'			oUser.Open sSql, Application("DSN"), 3, 1

			'If this is an INSERT then we need to get the new userid
'			if lcl_user_id = 0 then

				'Retrieve the userid that was just inserted
'				sSqlid = "SELECT IDENT_CURRENT('egov_users') as NewID"
'				oUser.Open sSqlid, Application("DSN"), 3, 1
'				lcl_identity = oUser.Fields("NewID").value

'				lcl_userid = lcl_identity
'			end if

'			oUser.Close
'			set oUser = Nothing 
		Else
			lcl_userid = 0
		End If 

	end if

	ProcessRecords = lcl_userid

end function

'------------------------------------------------------------------------------
sub UpdateEgovUsers(p_userid, p_confirmid)

  dim lcl_uid, lcl_cid

  lcl_uid = "NULL"
  lcl_cid = "NULL"

  if p_userid <> "" then
     lcl_uid = clng(p_userid)
  end if

  if p_confirmid <> "" then
     lcl_cid = p_confirmid
     lcl_cid = dbsafe(lcl_cid)
     lcl_cid = "'" & lcl_cid & "'"
  end if

	 sSQL = "UPDATE egov_users SET "
  sSQL = sSQL & " familyid = " & lcl_uid & ", "
  sSQL = sSQL & " subscription_confirmid = " & lcl_cid
  sSQL = sSQL & " WHERE userid = " & lcl_uid

 	RunSQLStatement sSQL

end sub

'------------------------------------------------------------------------------
sub AddFamilyMember(iBelongsToUserId, sFirstName, sLastName, sRelationship, sBirthDate)
 'This function adds family members to the egov_familymembers table
 	Dim sSql, oCmd

  lcl_firstname       = dbready_string(sFirstName,lcl_length_firstname)
  lcl_lastname        = dbready_string(sLastName,lcl_length_lastname)
  lcl_relationship    = dbready_string(sRelationship,lcl_length_relationship)
  lcl_birthdate       = "NULL"
  lcl_belongstouserid = 0

  if sBirthDate <> "" AND NOT isnull(sBirthday) then
     if dbready_date(sBirthDate) then
        lcl_birthdate = "'" & sBirthDate & "'"
     end if
  end if

  if dbready_number(iBelongsToUserID) then
     lcl_belongstouserid = iBelongsToUserID
  end if 

  sSQL = "INSERT INTO egov_familymembers (firstname, lastname, birthdate, belongstouserid, relationship, userid) VALUES ("
  sSQL = sSQL & "'" & lcl_firstname       & "', "
  sSQL = sSQL & "'" & lcl_lastname        & "', "
  sSQL = sSQL &       lcl_birthdate       & ", "
  sSQL = sSQL &       lcl_belongstouserid & ", "
  sSQL = sSQL & "'" & lcl_relationship    & "', "
  sSQL = sSQL &       lcl_belongstouserid
  sSQL = sSQL & ")"

  set rs = Server.CreateObject("ADODB.Recordset")
  rs.Open sSQL, Application("DSN"), 0, 1

  set rs = nothing

end sub

'------------------------------------------------------------------------------
function check_for_jobbid_categories(p_list_type, p_orgid )

  sSQL = "SELECT count(distributionlistid) AS total_count "
  sSQL = sSQL & " FROM egov_class_distributionlist "
  sSQL = sSQL & " WHERE orgid = " & p_orgid

  if p_list_type = "" then
     sSQL = sSQL & " AND (distributionlisttype = '' OR distributionlisttype IS NULL) "
  else
     sSQL = sSQL & " AND UPPER(distributionlisttype) = '" & UCASE(p_list_type) & "' "
  end if

  sSQL = sSQL & " AND distributionlistdisplay = 1 "

  set oPostingCategories = Server.CreateObject("ADODB.Recordset")
  oPostingCategories.Open sSQL, Application("DSN"), 0, 1

  if not oPostingCategories.eof then
     if oPostingCategories("total_count") > 0 then
        lcl_return = "Y"
     else
        lcl_return = "N"
     end if
  else
     lcl_return = "N"
  end if

  oPostingCategories.close
  set oPostingCategories = nothing

  check_for_jobbid_categories = lcl_return

end function

'------------------------------------------------------------------------------
function GetDistributionListName( iList )

	sSQL = "SELECT distributionlistname FROM egov_class_distributionlist WHERE distributionlistid = " & iList 

	set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	if not oList.eof then
    GetDistributionListName = "<strong>" & oList("distributionlistname") & "</strong><br />" & vbcrlf
 else
    GetDistributionListName = "" 
 end if

	oList.close 
	set oList = nothing

end function

'------------------------------------------------------------------------------
function IsMember( iUserID, listid )
  lcl_return = False

 	sSQL = "SELECT * "
  sSQL = sSQL & " FROM egov_class_distributionlist_to_user "
  sSQL = sSQL & " WHERE userid = " & iUserID
  sSQL = sSQL & " AND distributionlistid = " & listid

  set oList = Server.CreateObject("ADODB.Recordset")
  oList.Open sSQL, Application("DSN"), 0, 1

 	if not oList.EOF then
 		  lcl_return = True
 	end if

 	oList.close 
 	set oList = nothing

  IsMember = lcl_return

end function

'------------------------------------------------------------------------------
function createConfirmID(p_orgid, p_userid)

  lcl_return = ""
  lcl_year   = year(now())
  lcl_month  = month(now())
  lcl_day    = day(now())

  lcl_return = lcl_year & lcl_month & lcl_day & p_orgid & p_userid

  createConfirmID = lcl_return

end function

'------------------------------------------------------------------------------
function getOrgName(p_orgid)
  lcl_return = ""

  if p_orgid <> "" then
     sSQL = "SELECT orgname FROM organizations WHERE orgid = " & p_orgid

     set oGetOrgName = Server.CreateObject("ADODB.Recordset")
     oGetOrgName.Open sSQL, Application("DSN"), 3, 1

     if not oGetOrgName.eof then
        lcl_return = oGetOrgName("orgname")
     end if
				
     oGetOrgName.close
     set oGetOrgName = nothing
  end if

  getOrgName = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(p_value)
  lcl_return = ""

  if p_value <> "" then
     lcl_return = p_value
     lcl_return = replace(lcl_return,"'","''")
  end if

  dbsafe = lcl_return

end function

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

  set oDTB = Server.CreateObject("ADODB.Recordset")
  oDTB.Open sSQL, Application("DSN"), 0, 1

  set oDTB = nothing

end sub
%>
