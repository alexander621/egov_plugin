<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/inc_recordtoken.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="class/classFamily.asp" //-->
<!-- #include file="../egovlink300_global/includes/inc_passencryption.asp" //-->
<%	
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: register.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Citizen Registration.
'
' MODIFICATION HISTORY
' 1.1	12/26/2006	Steve Loar	- Changes for Menlo Park Project
' 1.2	08/15/2007	Steve Loar	- Changes to block spam attacks
' 1.3	10/26/2007	Steve Loar	- Added large address list selection and popup
' 2.0	10/02/2008	Steve Loar	- Changed way of saving to be more robust vs attacks
' 2.1	10/14/08	David Boyer - Added "fromPostings" request.querystring parameter
' 2.3	04/13/2010	Steve Loar - Changes to require address for Bullhead City
' 2.4	02/22/2011	Steve Loar - Making city, state and zip required optionally
' 2.5	08/23/2011	Steve Loar - Modify spam catch to catch zip codes of 123456
' 2.6	10/04/2011	Steve Loar - Added gender selection pick
' 2.7   2014-06-11  Jerry Felix - revised the email regex to be more permissive for new TLDs
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'PageDisplayCheck "registration", "", iUserId  ' In Common.asp	

assistant = "alexa"
if request.form("googleauth") = "true" then assistant = "googleauth"


 iUserId          = 0
 set oFamily      = New classFamily 
 bAddressRequired = False 

 If request.servervariables("request_method") = "POST" Then 
  	'a spam catch
	   If Trim(LCase((request("egov_users_useremail")))) = "admin@sexportal.com" Then 
     		response.redirect "register_none.asp#2"
   	End If 

   'Check the spam catching field to see if it has spam in it
    If request("subjecttext") <> "" Then 
		    'Commented out "sendspam email" via Peter's request as Christina is getting flooded with these emails. 11/16/10 - David Boyer
    	 	'SendSpamFlag request("egov_users_useremail"), request("subjecttext"), iorgid
     		response.redirect "register_none.asp#3"
    End If 

   	If request("egov_users_userzip") = "123456" Then
	 	   'This was added 8/23/2011, by Steve Loar to catch spam entries that are getting through the earlier checks above.
    		'If necessary this could be expanded to catch invalid states, or first name == last name.
     		response.redirect "register_none.asp#4"
   	End If 

  	'Check to see if user exists
    bUserFound = checkUserExists_byEmail(iorgid, request("egov_users_useremail"))

    If bUserFound Then 
     		errormsg = ""
		     errormsg = errormsg & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""color:#ffff00; font-size:8pt; background-color:#ff0000; border: solid 1px #000000;"">"
     		errormsg = errormsg & "  <tr>"
     		errormsg = errormsg & "      <td align=""center"">"
     		errormsg = errormsg & "          <strong><font color=""#000000"">Your registration could not be completed: </font><br />"
     		errormsg = errormsg & "          This email address is already associated with another registration.</strong>"
     		errormsg = errormsg & "      </td>"
     		errormsg = errormsg & "  </tr>"
     		errormsg = errormsg & "</table>"
   	Else 
     	'Check if last name and address are already in the table
     		If CheckUserExists (iorgid, request("egov_users_userlname"), request("egov_users_useraddress"), iUserId ) Then 
     			 'update existing record
       			UpdateExistingUser( iUserId )
    						InsertMailLists iUserId
    			Else 
    					'Try and filter out spam attacks
    						If Not IsNumeric(request("egov_users_neighborhoodid")) Then 
    			 				'redirect them to the spam fallout page
    			  				response.redirect "register_none.asp#5"
    						End If 

    					'Try and filter out spam attacks by checking the email
    						If isValidEmail( request("egov_users_useremail") ) And CLng(Len(request("egov_users_userpassword"))) < CLng(51) Then 
   			  				'Add them to the egov_users table
    			  				'userid = ProcessRecords()
         						'response.write "userid: " &  userid & "<br />"
    			  				userid = CreateCitizenRegistration()
			  								response.cookies("userid") = userid

			  							'Insert into the Family Members table 
			  								oFamily.AddFamilyMember userid, dbsafe(request("egov_users_userfname")), dbsafe(request("egov_users_userlname")), "Yourself", "NULL"

			  							'Update their FamilyId in egov_Users 
			  								oFamily.UpdateFamilyId userid, userid, request("egov_users_relationshipid"), request("skip_neighborhoodid")

			  							'Add the subscriptions
			  								InsertMailLists userid
			  					Else 
 		  							'redirect them to the spam fallout page
			  								response.redirect "register_none.asp#" & isValidEmail(request("egov_users_useremail"))
			  					End If 
			  		End If 

    		'Take them back to where they came from
		if request.form("token") <> "" then
		       	'RECORD THE TOKEN AS AUTHENTICATED
			RecordToken request.form("token"), iorgid


			response.redirect "basic_login.asp"
		elseif request(assistant) = "true" then
			'GENERATE CODE
			GUID = RecordGUID(request.form("state"), iorgid)

			'REDIRECT USER TO URI
				if assistant = "alexa" then
					response.redirect request.form("redirect_uri") & "#state=" & request.form("state") & "&token_type=Bearer&access_token=" & GUID
				else
					'response.redirect request.form("redirect_uri") & "#access_token=" & GUID & "&token_type=bearer&state=" & request.form("state")
					response.status = "302 Found"
					response.addheader "Location", request.form("redirect_uri") & "#access_token=" & GUID & "&token_type=bearer&state=" & request.form("state")
					response.end
				end if
     		elseIf session("RedirectPage") <> "" Then 
       			sRedirect = session("RedirectPage")
       			session("RedirectPage") = ""
			'response.redirect "default.asp?test=true"
       			response.redirect sRedirect
     		Else 
       			response.redirect GetEGovDefaultPage( iorgid )
     		End If 
    End If 
 End If 
%>
<html>
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	 <title>E-Gov Services <%=sOrgName%> - Registration</title>

	 <link rel="stylesheet" type="text/css" href="css/styles.css" />
	 <link rel="stylesheet" type="text/css" href="global.css" />
	 <link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

<style type="text/css">
  fieldset {
     border-radius: 6px;
  }

.address_fieldset {
   border:                1pt solid #808080;
   border-radius:         5px;
   -moz-border-radius:    5px;
   -webkit-border-radius: 5px;
}

#validaddresslist {
   border:                1pt solid #c0c0c0;
   border-radius:         6px;
  	-moz-border-radius:    6px;
   -webkit-border-radius: 6px;
   background-color:   #efefef;
   margin-top:         4px;
}

#validaddresslist legend {
   border:           1pt solid #c0c0c0;
   border-radius:    4px;
  	-moz-border-radius:    4px;
   -webkit-border-radius: 4px;
   background-color: #ffffff;
   color:            #ff0000;
   padding-left:     4px;
   padding-right:    4px;
}

div#addresspicklist {
  border-radius: 6px;
   -moz-border-radius:    6px;
   -webkit-border-radius: 5px;
}

.maintain_url {
   border:           1pt solid #000000;
   background-color: #c0c0c0;
   padding:          4px;
   color:            #000000;
   font-size:        10pt;
   display:          none;
}

.url_displaytext {
   font-size: 10pt;
   color:     #000000;
}

#screenMsg {
   text-align:  right;
   color:       #ff0000;
   font-size:   10pt;
   font-weight: bold;
}

.threehundredwide
{
	width:300px;
}
.twozeroeightwide
{
	width:208px;
}
</style>

</head>

<!--#Include file="include_top.asp"-->

<%

'BEGIN: Body Content ----------------------------------------------------------
RegisteredUserDisplay( "" )
%>
<!--#Include file="inc_register.asp"-->    

<!--#Include file="include_bottom.asp"-->    
<!--#Include file="includes\inc_dbfunction.asp"-->   

<%

'------------------------------------------------------------------------------
' Function GetDefaultRelationShipId( iOrgid )
'------------------------------------------------------------------------------
Function GetDefaultRelationShipId( ByVal iOrgid )
	Dim sSql, oRs

	sSql = "SELECT relationshipid FROM egov_familymember_relationships "
	sSql = sSql & "WHERE orgid = " & iorgid & " AND isdefault = 1"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetDefaultRelationShipId = oRs("relationshipid") 
	Else
		GetDefaultRelationShipId = 0 
	End if
	
	oRs.Close
	Set oRs = Nothing
	
End Function 


'------------------------------------------------------------------------------
' void UpdateExistingUser iUserId 
'------------------------------------------------------------------------------
Sub UpdateExistingUser( ByVal iUserId )
	Dim sSql

	sSql = "UPDATE egov_users SET "
	sSql = sSql & "userfname = '"           & dbsafe(request("egov_users_userfname"))    & "', "
	sSql = sSql & "userstate = '"           & request("egov_users_userstate")            & "', "
	sSql = sSql & "usercity = '"            & request("egov_users_usercity")             & "', "
	sSql = sSql & "userzip = '"             & request("egov_users_userzip")              & "', "
	sSql = sSql & "useremail = '"           & request("egov_users_useremail")            & "', "
	sSql = sSql & "userhomephone = '"       & request("egov_users_userhomephone")        & "', "
	sSql = sSql & "userworkphone = '"       & request("egov_users_userworkphone")        & "', "
	sSql = sSql & "userbusinessname = '"    & request("egov_users_userbusinessname")     & "', "
	sSql = sSql & "userbusinessaddress = '" & request("egov_users_userbusinessaddress")  & "', "
	sSql = sSql & "userpassword = '"        & dbsafe(request("egov_users_userpassword")) & "', "
	sSql = sSql & "residenttype = '"        & request("egov_users_residenttype")         & "' "
	sSql = sSql & " WHERE userid = " & iUserId
	'	response.write sSql

	RunSQLStatement sSql		' In common.asp

End Sub 


'------------------------------------------------------------------------------
' void InsertMailLists iuserid 
'------------------------------------------------------------------------------
Sub InsertMailLists( ByVal iuserid )
	Dim sSql

	sSql = "DELETE FROM egov_class_distributionlist_to_user WHERE userid = " & iuserid 
	RunSQLStatement sSql		' In common.asp

	' Insert subscriptions
	For Each list In request("maillist")
		If IsNumeric(list) Then 
			sSql = "INSERT INTO egov_class_distributionlist_to_user ( userid, distributionlistid ) VALUES ( " & iuserid & ", " & list & " )"
			RunSQLStatement sSql		' In common.asp
		End If 
	Next	

End Sub




'------------------------------------------------------------------------------
' integer CreateCitizenRegistration( )
'------------------------------------------------------------------------------
Function CreateCitizenRegistration( )
	Dim sSql, bValid, sUserfname, sUserlname, sUserstreetnumber, sUserunit, sUseraddress, bTestValid
	Dim sUsercity, sUserstate, sUserzip, oRE, sUseremail, sUserhomephone, sUsercell, sUserworkphone
	Dim sUserfax, sUserbusinessname, sUserbusinessaddress, sEmergencycontact, sUserpassword
	Dim sResidenttype, sUserbusinessnumber, iRelationshipid, iNeighborhoodid, sEmergencyphone
	Dim sGender

	Set oRE = New RegExp
	oRE.IgnoreCase = False    
	oRE.Global = False 

	' Format and Validate before inserting to filter out the spammers and attackers 
	bValid = True 

	sUserfname = "'" & dbready_string(request("egov_users_userfname"),50) & "'"

	sUserlname = "'" & dbready_string(request("egov_users_userlname"),50) & "'"

	If request("egov_users_gender") <> "M" And request("egov_users_gender") <> "F" Then
		sGender = "NULL"
	Else
		sGender = "'" & dbready_string(request("egov_users_gender"),1) & "'"
	End If 

	If request("residentstreetnumber") <> "" Then
		sUserstreetnumber = "'" & dbready_string(request("residentstreetnumber"),10) & "'"
	Else
		sUserstreetnumber = "NULL"
	End If 

	If request("egov_users_userunit") <> "" Then
		If clng(InStr(request("egov_users_userunit"),"http:")) > clng(0) Then 
			bValid = False 
		Else
			sUserunit = "'" & dbready_string(request("egov_users_userunit"),10) & "'"
		End If 
	Else
		sUserunit = "NULL"
	End If

	If request("egov_users_useraddress") <> "" Then 
		sUseraddress = "'" & dbready_string(request("egov_users_useraddress"),255) & "'"
	Else
		sUseraddress = "NULL"
	End If 

	If request("egov_users_usercity") <> "" Then
		sUsercity = "'" & dbready_string(request("egov_users_usercity"),50) & "'"
	Else
		sUsercity = "NULL"
	End If 

	If request("egov_users_userstate") <> "" Then
		sUserstate = UCase(dbready_string(request("egov_users_userstate"),2))
		' Spammers put a foreign country in as the state, so this will trip them up.
		'If StateNotValid( sUserstate ) Then
		'	bValid = False 
		'	sUserstate = "'" & sUserstate & "'"
		'Else
			sUserstate = "'" & sUserstate & "'"
		'End If 
	Else
		sUserstate = "NULL"
	End If 

	If request("egov_users_userzip") <> "" Then 
		sUserzip = "'" & dbready_string(request("egov_users_userzip"),10) & "'"
		' Check for numbers and a dash only - Spammers may put too many digits and no dash
		'If clng(InStr(sUserzip, "-")) = clng(0) Then
	    '    oRE.Pattern = "^'\d{5}'$"
        'Else
		'    oRE.Pattern = "^'\d{5}-\d{4}'$"
        'End If
		'bTestValid = oRE.Test(sUserzip)
		'response.write "sUserzip " & sUserzip & " " & bValid & " pattern: " & oRE.Pattern & "<br />"
		'If bTestValid = False Then
		'	bValid = False
		'End If 
	Else
		sUserzip = "NULL"
	End If 

	If request("egov_users_useremail") <> "" Then 
		sUseremail = "'" & dbready_string(request("egov_users_useremail"),512) & "'"
	Else
		bValid = False    ' They have to have an email
	End If 

	If request("egov_users_userhomephone") <> "" Then
		sUserhomephone = "'" & dbready_string(request("egov_users_userhomephone"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		bTestValid = oRE.Test(sUserhomephone)
		If bTestValid = False Then
			bValid = False
		End If 
	Else
		sUserhomephone = "NULL"
	End If 

	If request("egov_users_usercell") <> "" Then
		sUsercell = "'" & dbready_string(request("egov_users_userhomephone"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		bTestValid = oRE.Test(sUsercell)
		If bTestValid = False Then
			bValid = False
		End If 
	Else
		sUsercell = "NULL"
	End If 

	If request("egov_users_userworkphone") <> "" Then
		sUserworkphone = "'" & dbready_string(request("egov_users_userworkphone"),14) & "'"
		oRE.Pattern = "^'\d{10}\d{0,4}'$"
		bTestValid = oRE.Test(sUserworkphone)
		If bTestValid = False Then
			bValid = False
		End If 
	Else
		sUserworkphone = "NULL"
	End If 

	If request("egov_users_userfax") <> "" Then
		sUserfax = "'" & dbready_string(request("egov_users_userfax"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		bTestValid = oRE.Test(sUserfax)
		If bTestValid = False Then
			bValid = False
		End If 
	Else
		sUserfax = "NULL"
	End If 

	If request("egov_users_userbusinessname") <> "" Then
		sUserbusinessname = "'" & dbready_string(request("egov_users_userbusinessname"),100) & "'"
	Else
		sUserbusinessname = "NULL"
	End If 

	If request("egov_users_userbusinessaddress") <> "" Then 
		sUserbusinessaddress = "'" & dbready_string(request("egov_users_userbusinessaddress"),255) & "'"
	Else
		sUserbusinessaddress = "NULL"
	End If 

	If request("egov_users_userpassword") <> "" Then
		If request("skip_userpassword2") <> request("egov_users_userpassword") Then 
			bValid = False
		Else 
			sUserpassword = "'" & createHashedPassword(request("egov_users_userpassword")) & "'"
		End If 
	Else
		' passwords are required
		bValid = False 
	End If

	If request("egov_users_residenttype") <> "" Then 
		sResidenttype = "'" & dbready_string(request("egov_users_residenttype"),1) & "'"

		if sResidenttype = "'N'" and sUserzip = "'94025'" and lcase(sUsercity) = "'menlo park'" then sResidenttype = "'U'"
	Else
		' This will always have a value
		bValid = False 
	End If

	If request("egov_users_userbusinessnumber") <> "" Then 
		sUserbusinessnumber = "'" & dbready_string(request("egov_users_userbusinessnumber"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		bTestValid = oRE.Test(sUserbusinessnumber)
		If bTestValid = False Then
			bValid = False
		End If 
	Else
		sUserbusinessnumber = "NULL"
	End If

	If request("egov_users_relationshipid") <> "" Then
		If dbready_number( request("egov_users_relationshipid") ) Then
			iRelationshipid = CLng(request("egov_users_relationshipid"))
		Else
			bValid = False 
		End If 
	Else
		' This will always have a value
		bValid = False 
	End If 

	If request("egov_users_neighborhoodid") <> "" Then
		If dbready_number( request("egov_users_neighborhoodid") ) Then
			iNeighborhoodid = CLng(request("egov_users_neighborhoodid"))
		Else
			bValid = False 
		End If 
	Else
		' This will always have a value
		bValid = False 
	End If 

	If request("egov_users_emergencycontact") <> "" Then
		If clng(InStr(request("egov_users_emergencycontact"),"http://")) > clng(0) Then 
			bValid = False 
		Else
			sEmergencycontact = "'" & dbready_string(request("egov_users_emergencycontact"),100) & "'"
		End If 
	Else
		sEmergencycontact = "NULL"
	End If

	If request("egov_users_emergencyphone") <> "" Then
		sEmergencyphone = "'" & dbready_string(request("egov_users_emergencyphone"),10) & "'"
		oRE.Pattern = "^'\d{10}'$"
		bTestValid = oRE.Test(sEmergencyphone)
		If bTestValid = False Then
			bValid = False
		End If 
	Else
		sEmergencyphone = "NULL"
	End If 

	sIsOnDoNotKnockList_peddlers   = "0"
	sIsOnDoNotKnockList_solicitors = "0"
	sIsDoNotKnockVendor_peddlers   = "0"
	sIsDoNotKnockVendor_solicitors = "0"

	If request("isOnDoNotKnockList_peddlers") = "on" Then 
		sIsOnDoNotKnockList_peddlers = "1"
	End If 

	If request("isOnDoNotKnockList_solicitors") = "on" Then 
		sIsOnDoNotKnockList_solicitors = "1"
	End If 

	'if request("isDoNotKnockVendor_peddlers") = "on" then
	'   sIsDoNotKnockVendor_peddlers = "1"
	'end if

	'if request("isDoNotKnockVendor_solicitors") = "on" then
	'   sIsDoNotKnockVendor_solicitors = "1"
	'end if

	Set oRE = Nothing  

	If Not bValid Then
		response.redirect "register_none.asp#1"
		'response.write bValid
	Else 
		' compose the insert statement
		sSql = "INSERT INTO egov_users ("
		sSql = sSql & "orgid,"
		sSql = sSql & "userfname,"
		sSql = sSql & "userlname,"
		sSql = sSql & "userstreetnumber,"
		sSql = sSql & "userunit,"
		sSql = sSql & "useraddress,"
		sSql = sSql & "usercity,"
		sSql = sSql & "userstate,"
		sSql = sSql & "userzip,"
		sSql = sSql & "useremail,"
		sSql = sSql & "userhomephone,"
		sSql = sSql & "usercell,"
		sSql = sSql & "userworkphone,"
		sSql = sSql & "userfax,"
		sSql = sSql & "userbusinessname,"
		sSql = sSql & "password,"
		sSql = sSql & "userregistered,"
		sSql = sSql & "residenttype,"
		sSql = sSql & "userbusinessnumber,"
		sSql = sSql & "userbusinessaddress,"
		sSql = sSql & "relationshipid,"
		sSql = sSql & "neighborhoodid,"
		sSql = sSql & "emergencycontact,"
		sSql = sSql & "emergencyphone,"
		sSql = sSql & "headofhousehold,"
		sSql = sSql & "isOnDoNotKnockList_peddlers,"
		sSql = sSql & "isOnDoNotKnockList_solicitors,"
		sSql = sSql & "isDoNotKnockVendor_peddlers, "
		sSql = sSql & "isDoNotKnockVendor_solicitors, "
		sSql = sSql & "gender "
		sSql = sSql & ") VALUES ( "
		sSql = sSql & iorgid               & ", "
		sSql = sSql & sUserfname           & ", "
		sSql = sSql & sUserlname           & ", "
		sSql = sSql & sUserstreetnumber    & ", "
		sSql = sSql & sUserunit            & ", "
		sSql = sSql & sUseraddress         & ", "
		sSql = sSql & sUsercity            & ", "
		sSql = sSql & sUserstate           & ", "
		sSql = sSql & sUserzip             & ", "
		sSql = sSql & sUseremail           & ", "
		sSql = sSql & sUserhomephone       & ", "
		sSql = sSql & sUsercell            & ", "
		sSql = sSql & sUserworkphone       & ", "
		sSql = sSql & sUserfax             & ", "
		sSql = sSql & sUserbusinessname    & ", "
		sSql = sSql & sUserpassword        & ", "
		sSql = sSql & "1, "
		sSql = sSql & sResidenttype        & ", "
		sSql = sSql & sUserbusinessnumber  & ", "
		sSql = sSql & sUserbusinessaddress & ", "
		sSql = sSql & iRelationshipid      & ", "
		sSql = sSql & iNeighborhoodid      & ", "
		sSql = sSql & sEmergencycontact    & ", "
		sSql = sSql & sEmergencyphone      & ", "
		sSql = sSql & "1, "
		sSql = sSql & sIsOnDoNotKnockList_peddlers   & ", "
		sSql = sSql & sIsOnDoNotKnockList_solicitors & ", "
		sSql = sSql & sIsDoNotKnockVendor_peddlers   & ", "
		sSql = sSql & sIsDoNotKnockVendor_solicitors & ", "
		sSql = sSql & sGender
		sSql = sSql & " )"

		'response.write sSql

		CreateCitizenRegistration = RunIdentityInsert( sSql )	
	End If 
	
End Function 


'------------------------------------------------------------------------------
' void SendSpamFlag sFromEmail, sTextinput, sOrgID
'------------------------------------------------------------------------------
Sub SendSpamFlag( ByVal sFromEmail, ByVal sTextinput, ByVal sOrgID )
	Dim sOrgName, lcl_from, lcl_to, lcl_subject, sMsgBody

	'Setup the email variables
	sOrgName = getOrgName(sOrgID)
	lcl_from = sOrgName & " E-GOV WEBSITE <noreply@eclink.com>"
	lcl_to = "egovsupport@eclink.com"
	lcl_subject = sOrgName & " - E-Gov (Possible Subscription Spam Submission)"

	'Build the message
	sMsgBody = "Possible CITIZEN REGISTRATION spam submitted to " & sOrgName & ".<br />"
	sMsgBody = sMsgBody & "<p>Citizen Email: "   & sFromEmail  & "<br />"
	sMsgBody = sMsgBody & "Hidden field contains: " & sTextinput  & "</p>"

	'Send the email
	'sendEmail lcl_from, lcl_to, "", lcl_subject, sMsgBody, "", "Y"

End Sub 


'------------------------------------------------------------------------------
' string getOrgName( p_orgid )
'------------------------------------------------------------------------------
Function getOrgName( ByVal p_orgid )
	Dim sSql, oRs

	If p_orgid <> "" Then 
		sSql = "SELECT orgname FROM organizations WHERE orgid = " & p_orgid

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then 
			getOrgName = oRs("orgname")
		Else
			' this is a problem
			getOrgName = ""
		End If 

		oRs.Close
		Set oRs = Nothing 
	Else
		getOrgName = ""
	End If 

End Function 


'------------------------------------------------------------------------------
' boolean CheckUserExists( iorgid, sUserlname, sUseraddress, iUserId )
'------------------------------------------------------------------------------
Function CheckUserExists( ByVal iorgid, ByVal sUserlname, ByVal sUseraddress, ByRef iUserId )
	Dim sSql, oRs

	sUseraddress = Trim(UCase(sUseraddress))
	sUserlname = Trim(UCase(sUserlname))

	sSql = "SELECT userid FROM egov_users "
	sSql = sSql & " WHERE orgid = " & iorgid
	sSql = sSql & " AND UPPER(userlname) = '" & dbsafe(sUserlname) & "' "
	sSql = sSql & " AND upper(useraddress) = '" & dbsafe(sUseraddress) & "' "
	sSql = sSql & " AND useremail = 'webmaster@ci.montgomery.oh.us' "

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iUserId = oRs("userid")
		CheckUserExists = True 
	Else
		CheckUserExists = False 
	End if

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' boolean checkUserExists_byEmail( p_orgid, p_useremail )
'------------------------------------------------------------------------------
Function checkUserExists_byEmail( ByVal p_orgid, ByVal p_useremail )
	Dim sSql, oRs, lcl_useremail, lcl_orgid

	lcl_useremail = "''"
	lcl_orgid = CLng(p_orgid)

	If p_useremail <> "" Then 
		lcl_useremail = "'" & dbsafe(p_useremail) & "'"
	End if 

	sSql = "SELECT userid FROM egov_users "
	sSql = sSql & " WHERE isdeleted = 0 AND userregistered = 1 AND useremail = " & lcl_useremail
	sSql = sSql & " AND orgid = " & lcl_orgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		checkUserExists_byEmail = True
	Else
		checkUserExists_byEmail = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>
