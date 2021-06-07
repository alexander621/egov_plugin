<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../include_top_functions.asp"-->
<!-- #include file="../class/classOrganization.asp"-->
<!-- #include file="subscribe_global_functions.asp" //-->
<%
  lcl_action              = ""
  lcl_cid                 = ""
  lcl_subscribe_confirmid = ""
  session("useremail")    = ""
  session("userpassword") = request("userpassword")

 'Check for org feature
  lcl_orghasfeature_subscriptions_checkforuseraddress = orghasfeature(iorgid, "subscriptions_checkforuseraddress")

 'Determine if user is coming from "unsubscribe url" found within the new subscription email signup.
  if request("c") <> "" then
     lcl_subscribe_confirmid = request("c")
  end if

  if lcl_subscribe_confirmid <> "" then
    'If a user exists then verify the password entered.
     checkUserExists "", _
                     "", _
                     "", _
                     lcl_subscribe_confirmid, _
                     lcl_user_id, _
                     lcl_user_email, _
                     lcl_user_password, _
                     lcl_user_address

     if lcl_user_id <> "" then
        lcl_cid = lcl_subscribe_confirmid

       'Get the egov url for the org
        sSQL = "SELECT OrgEgovWebsiteURL "
        sSQL = sSQL & " FROM organizations "
        sSQL = sSQL & " WHERE orgid IN (select orgid "
        sSQL = sSQL &                 " from egov_users "
        sSQL = sSQL &                 " where subscription_confirmid = '" & lcl_cid & "') "

        set oGetOrgURL = Server.CreateObject("ADODB.Recordset")
        oGetOrgURL.Open sSQL, Application("DSN"), 0, 1

        if not oGetOrgURL.eof then
           lcl_org_url = replace(oGetOrgURL("orgEgovWebsiteURL"), "www", Application("subscriptions_environment")) & "/subscriptions/subscribe.asp?c=" & lcl_cid

           response.redirect lcl_org_url
        end if

        oGetOrgURL.close
        set oGetOrgURL = nothing
     end if
  else
     if request("doaction") <> "" then
        lcl_action = UCASE(request("doaction"))
     end if

    'Determine if the user is setting up a new subscription(s) or searching for existing subscriptions.
     lcl_useremail = Trim(request("useremail"))

     if lcl_action = "SUBSCRIBE" then
        lcl_userid  = ""
        lcl_success = ""

        'Check the spam catching field to see if it has spam in it
        if request("subjecttext") <> "" then
           'SendSpamFlag request("useremail"), request("userpassword"), request("subjecttext"), iorgid, request.servervariables("remote_addr")

           response.redirect "subscription_none.asp"
    	end if

        'If a user exists then verify the password entered.
        checkUserExists "", _
                        lcl_useremail, _
                        iorgid, _
                        "", _
                        lcl_user_id, _
                        lcl_user_email, _
                        lcl_user_password, _
                        lcl_user_address

       	if lcl_user_id <> "" then

			'Sets parameter to be used in response.redirect
			lcl_userid = lcl_user_id

			'check their password
			'if lcl_user_password = request("userpassword") then
			if ValidateUser(request("userpassword"), lcl_user_password)  then

				'Clear existing values
				ClearMailList lcl_user_id

				'Create the subscription record(s)
				for each list in request("maillist")
					InsertMailList lcl_user_id, list
				next

				'Log the user in
				response.cookies("userid") = lcl_user_id
				lcl_userid                 = lcl_user_id

				'Build the return URL
				lcl_success = "SU"
			else
				'Password does not match
				lcl_success             = "EXISTS_BAD_PWD"
				session("useremail")    = lcl_useremail
				'session("userpassword") = request("userpassword")
			end if
        else
         	if request("userpassword") <> "" then

				'Add them to the egov_users table if they have a valid email etc.
				subscriberuserid = ProcessRecords(iorgid, Replace(request("useremail"),"'",""), request("userpassword"), request("residenttype"), request("userregistered"), request("headofhousehold"))

				If CLng(subscriberuserid) > CLng(0) Then 
					'Create the confirmation id
					lcl_confirmationid = createConfirmID(iorgid, subscriberuserid)

					'Update the egov_users for familyid
					UpdateEgovUsers subscriberuserid, lcl_confirmationid

					'Insert into the Family Members table
					AddFamilyMember subscriberuserid, "", "", "Yourself", ""

					'Add the subscriptions
					'InsertTempMailLists subscriberuserid
					for each list in request("maillist")
						  InsertMailList subscriberuserid, list
					next

					'Send them a confirmation email
					SendConfirmation request("useremail"), subscriberuserid, request("userpassword"), lcl_confirmationid

					'Log the user in
					response.cookies("userid") = subscriberuserid
					lcl_userid                 = subscriberuserid

					'Build the return URL
					lcl_success = "SA"
				Else
					'This will catch some more spam entries 
					lcl_success = "NOT_EXISTS"
				End If 
         	else
         		'This will catch spam entries that have a blank password
				lcl_success = "REQUIRED_PWD"
        	end if

            'response.redirect "subscribe.asp?success=" & lcl_success
       	end if

     elseif lcl_action = "UNSUBSCRIBE" OR lcl_action = "MANAGE SUBSCRIPTIONS" then

        if request("subscription_confirmid") <> "" then
           lcl_subscribe_confirmid = request("subscription_confirmid")
        end if

       'If a user exists then verify the password entered.
        checkUserExists "", _
                        "", _
                        "", _
                        lcl_subscribe_confirmid, _
                        lcl_user_id, _
                        lcl_user_email, _
                        lcl_user_password, _
                        lcl_user_address

       	if lcl_user_id <> "" then

          'Clear existing values
           if lcl_action = "UNSUBSCRIBE" then
              ClearMailList lcl_user_id
              lcl_success = "UNSUBSCRIBED"
           else
              session("useremail") = lcl_user_email
              lcl_success          = "MANAGE"
           end if
        end if

     else  'Find subscriptions

        checkUserExists "", _
                        lcl_useremail, _
                        iorgid, _
                        "", _
                        lcl_user_id, _
                        lcl_user_email, _
                        lcl_user_password, _
                        lcl_user_address

      		if lcl_user_id <> "" then

          'Sets parameter to be used in response.redirect
        			lcl_userid = lcl_user_id

        			'if lcl_user_password <> request("userpassword") then
        			if not ValidateUser(request("userpassword"), lcl_user_password) then
         				'Password does not match
              lcl_success             = "BAD_PWD"
              session("useremail")    = lcl_useremail
              'session("userpassword") = request("userpassword")
       	   			'sMsg = "The password you entered is incorrect."
        			else
          			'Found and password matches
          				response.cookies("userid") = lcl_userid
      		   end if
       	else
           lcl_success             = "NOT_EXISTS"
           session("useremail")    = lcl_useremail
           'session("userpassword") = request("userpassword")
        			'sMsg = "We do not have an account with this email address."
       	end if

     end if
  end if

 'Check for url parameters
  lcl_url_params = ""

  if lcl_userid <> "" then
     if lcl_url_params = "" then
        lcl_url_params = "?u=" & lcl_userid
     else
        lcl_url_params = lcl_url_params & "&u=" & lcl_userid
     end if
  end if

  if lcl_cid <> "" then
     if lcl_url_params = "" then
        lcl_url_params = "?c=" & lcl_cid
     else
        lcl_url_params = lcl_url_params & "&c=" & lcl_cid
     end if
  end if

 'Determine if we are to check for an address.
 'If "yes" and no address exists then redirect user to "manage account" screen
 'If "no" then return a success message
 '  SU_NA = Successfully Updated - No Address
 '  SA_NA = Successfully Added   - No Address
  if lcl_orghasfeature_subscriptions_checkforuseraddress AND lcl_action = "SUBSCRIBE" then
     if lcl_user_address <> "" then
        lcl_return_url = "subscribe.asp"
     else
        lcl_return_url = "../manage_account.asp"
        lcl_success    = lcl_success & "_NA"
     end if
  else
     lcl_return_url = "subscribe.asp"
  end if

  if lcl_success <> "" then
     if lcl_url_params = "" then
        lcl_url_params = "?success=" & lcl_success
     else
        lcl_url_params = lcl_url_params & "&success=" & lcl_success
     end if
  end if

	LogThePage


  response.redirect lcl_return_url & lcl_url_params

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
	sSql = sSql & " servername, remoteaddress, requestmethod, orgid, userid, username, sectionid, documenttitle, useragent, useragentgroup, requestformcollection ) VALUES ( "
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
	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
		sSql = sSql & request.cookies("userid") & ", "
	Else
		sSql = sSql & "NULL, "
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

	sSql = sSql & "'" & DBsafe(GetRequestformInformation()) & "'"

	sSql = sSql & " )"
	'response.write sSql
	'response.end

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql

	session("sSql") = sSql
	oCmd.Execute
	session("sSql") = ""

	Set oCmd = Nothing

End Sub 
'------------------------------------------------------------------------------
Function GetUserAgentGroup( ByVal sUserAgent )
	Dim sSql, oRs, sUserAgentGroup

	sUserAgentGroup = GetUntrackedUserAgentGroup()

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 0 AND isactive = 1 ORDER BY checkorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If clng(InStr( 1, sUserAgent, LCase(oRs("useragentgroup")), 1 )) > clng(0) Then
			sUserAgentGroup = oRs("useragentgroup")
			Exit Do 
		End If 
		oRs.MoveNext
	Loop 
	
	oRs.Close
	Set oRs = Nothing 
	
	GetUserAgentGroup = sUserAgentGroup

End Function 


'------------------------------------------------------------------------------
Function GetUntrackedUserAgentGroup( )
	Dim sSql, oRs

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 1"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetUntrackedUserAgentGroup = oRs("useragentgroup")
	Else
		GetUntrackedUserAgentGroup = "untracked"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 

'--------------------------------------------------------------------------------------------------
' FUNCTION GETREQUESTFORMINFORMATION()
'--------------------------------------------------------------------------------------------------
Function GetRequestFormInformation()
	Dim sReturnValue, key
	
	sReturnValue = ""

	For each key in request.Form
		If key <> "accountnumber" And key <> "cvv2" Then 
			sReturnValue = sReturnValue & key & ":" & request.form(key) & "<br />" & vbcrlf
		End If 
	Next 
	
	GetRequestFormInformation = sReturnValue

End Function

%>
