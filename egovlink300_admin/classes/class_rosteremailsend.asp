<!-- #include file="../includes/common.asp" -->
<!-- #include file="class_global_functions.asp" -->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_rosteremailsend.asp
' AUTHOR: Steve Loar
' CREATED: 04/26/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module sends emails to a class/event roster
'
' MODIFICATION HISTORY
' 1.0 04/26/06	Steve Loar - Initial Version
' 1.1	05/03/07	Steve Loar - Changes for Menlo Park Project
' 1.2 2014-06-10 Jerry Felix - Changed line 26 (was #25) from cInt conversion to Int conversion - timeid has exceeded 32767
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
	Dim iClassId, bIsParent, sMessageBody, iTimeId, sClassName, sSubject, sFromEmail, sFromName, iSentCount, sWhere

	iClassId     = clng(request("classid"))
	iTimeId      = Int(request("timeid"))
	sMessageBody = request("messagebody")
	sSubject     = request("subject")
	sFromEmail   = request("fromemail")
	sFromName    = request("fromname")
	iSentCount   = 0
	sWhere       = ""

	select case clng(request("sendto"))
  	case 1
     sWhere = " and L.status = 'ACTIVE' "
 		case 2
		  	sWhere = " and L.status = 'WAITLIST' "
	 	case 3
  			sWhere = " and (L.status = 'ACTIVE' OR L.status = 'WAITLIST') "
	end select 

'Get isParent
	bIsParent = GetClassIsParent( iClassId )

'Get the class name
 'sClassName = GetClassName( iClassId )
	
	if bIsParent then
 		'send email to children
  		iSentCount = SendChildrenMail( iClassId, sMessageBody, sSubject, sFromEmail, sFromName, sWhere )
 else
 		'just send mail to the regular class/event attendees
		  iSentCount = SendClassMail( iClassId, iTimeId, sMessageBody, sSubject, sFromEmail, sFromName, sWhere )
 end if

	'response.redirect "class_rosteremailconfirmation.asp?classid=" & iClassId & "&timeid=" & iTimeId & "&sentcount=" & iSentCount
 response.redirect "view_roster.asp?classid=" & iClassID & "&timeid=" & iTimeID & "&sentcount=" & iSentCount & "&success=SS"

'------------------------------------------------------------------------------
function GetClassIsParent( iClassId )
 	Dim sSQL, oInfo

 	sSQL = "SELECT isparent FROM egov_class WHERE classid = " & iClassId
 	GetClassIsParent = False

 	set oInfo = Server.CreateObject("ADODB.Recordset")
 	oInfo.Open sSQL, Application("DSN"), 0, 1

 	if not oInfo.eof then
   		GetClassIsParent = oInfo("isparent")
 	end if

 	oInfo.close
 	set oInfo = nothing

end function

'------------------------------------------------------------------------------
function SendClassMail( iClassId, iTimeId, sMessageBody, sSubject, sFromEmail, sFromName, sWhere )
 	Dim sSQL, oEmails, iSentCount

	'There is One email shared by the family
 	iSentCount = 0
  lcl_where  = ""

 	if ClassRequiresRegistration( iClassId ) then
   		sSQL = "SELECT useremail, F.firstname, F.lastname, isnull(L.attendeeuserid,L.userid) as emailuserid "
     sSQL = sSQL & " FROM egov_class C, egov_class_list L, egov_users U, egov_familymembers F "

   		lcl_where = " AND L.familymemberid = F.familymemberid "

   		'sSQL = sSQL & " WHERE C.classid = L.classid "
   		'sSQL = sSQL & " AND L.userid = U.userid "
   		'sSQL = sSQL & sWhere 
   		'sSQL = sSQL & " AND L.familymemberid = F.familymemberid "
   		'sSQL = sSQL & " AND C.classid = " & iClassId
   		'sSQL = sSQL & " AND L.classtimeid = " & iTimeId
   		'sSQL = sSQL & " AND useremail not in ('webmaster@ci.montgomery.oh.us', 'customer_service@ci.montgomery.oh.us') "
   		'sSQL = sSQL & " AND U.emailnotavailable = 0 "
 	else ' Ticketed event
   		sSQL = "SELECT U.useremail, U.userfname as firstname, U.userlname as lastname, isnull(L.attendeeuserid,L.userid) as emailuserid "
   		sSQL = sSQL & " FROM egov_class C, egov_class_list L, egov_users U "
   		'sSQL = sSQL & " WHERE C.classid = L.classid "
   		'sSQL = sSQL & " AND L.userid = U.userid "
   		'sSQL = sSQL & sWhere 
   		'sSQL = sSQL & " AND C.classid = " & iClassId
   		'sSQL = sSQL & " AND L.classtimeid = " & iTimeId
   		'sSQL = sSQL & " AND useremail not in ('webmaster@ci.montgomery.oh.us', 'customer_service@ci.montgomery.oh.us') "
   		'sSQL = sSQL & " AND U.emailnotavailable = 0 "
 	end if

 	sSQL = sSQL & " WHERE C.classid = L.classid "
  sSQL = sSQL & " AND L.userid = U.userid "
  sSQL = sSQL & sWhere 

  if lcl_where <> "" then
     sSQL = sSQL & lcl_where
  end if

  sSQL = sSQL & " AND C.classid = " & iClassId
  sSQL = sSQL & " AND L.classtimeid = " & iTimeId
  sSQL = sSQL & " AND useremail not in ('webmaster@ci.montgomery.oh.us', 'customer_service@ci.montgomery.oh.us') "
  sSQL = sSQL & " AND U.emailnotavailable = 0 "

 	set oEmails = Server.CreateObject("ADODB.Recordset")
 	oEmails.Open sSQL, Application("DSN"), 0, 1

 	do while not oEmails.eof
    	if oEmails("useremail") <> "" then
     			subSendEmail oEmails("firstname") & " " & oEmails("lastname"), oEmails("emailuserid"), sFromName, sFromEmail, sSubject, sMessageBody , oEmails("useremail")
     			iSentCount = iSentCount + 1
   		end if

   		oEmails.movenext
 	loop

 	oEmails.close
 	set oEmails = nothing

 	SendClassMail = iSentCount

end function

'------------------------------------------------------------------------------
function SendChildrenMail( iParentClassId, sMessageBody, sSubject, sFromEmail, sFromName, sWhere )
 	Dim sSQL, oEmails, iSentCount

	'There is One email shared by the family
 	iSentCount = 0
  lcl_where  = ""

 	if ClassRequiresRegistration( iClassId ) then
   		sSQL = "SELECT isnull(useremail,'') as useremail, classname, F.firstname, F.lastname, isnull(L.attendeeuserid,L.userid) as emailuserid "
  		 sSQL = sSQL & " FROM egov_class C, egov_class_list L, egov_users U, egov_familymembers F "

   		lcl_where = " AND L.familymemberid = F.familymemberid "

		   'sSQL = sSQL & " WHERE C.classid = L.classid "
  		 'sSQL = sSQL & " AND L.userid = U.userid "
  		 'sSQL = sSQL & sWhere 
   		'sSQL = sSQL & " AND L.familymemberid = F.familymemberid "
   		'sSQL = sSQL & " AND C.parentclassid = " & iParentClassId
  		 'sSQL = sSQL & " AND useremail not in ('webmaster@ci.montgomery.oh.us', 'customer_service@ci.montgomery.oh.us') "
  		 'sSQL = sSQL & " AND U.emailnotavailable = 0 "
 	else  ' Ticketed Event
	   	sSQL = "SELECT isnull(useremail,'') as useremail, classname, U.userfname as firstname, F.userlname as lastname, isnull(L.attendeeuserid,L.userid) as emailuserid "
   		sSQL = sSQL & " FROM egov_class C, egov_class_list L, egov_users U "
   		'sSQL = sSQL & " WHERE C.classid = L.classid "
  		 'sSQL = sSQL & " AND L.userid = U.userid "
  		 'sSQL = sSQL & sWhere
   		'sSQL = sSQL & " AND C.parentclassid = " & iParentClassId
  		 'sSQL = sSQL & " AND useremail not in ('webmaster@ci.montgomery.oh.us', 'customer_service@ci.montgomery.oh.us') "
  		 'sSQL = sSQL & " AND U.emailnotavailable = 0 "
 	end if

  sSQL = sSQL & " WHERE C.classid = L.classid "
  sSQL = sSQL & " AND L.userid = U.userid "
  sSQL = sSQL & sWhere

  if lcl_where <> "" then
     sSQL = sSQL & lcl_where
  end if

  sSQL = sSQL & " AND C.parentclassid = " & iParentClassId
  sSQL = sSQL & " AND useremail not in ('webmaster@ci.montgomery.oh.us', 'customer_service@ci.montgomery.oh.us') "
  sSQL = sSQL & " AND U.emailnotavailable = 0 "

 	set oEmails = Server.CreateObject("ADODB.Recordset")
	 oEmails.Open sSQL, Application("DSN"), 0, 1

 	do while not oEmails.eof
	   	if oEmails("useremail") <> "" then
     			subSendEmail oEmails("firstname") & " " & oEmails("lastname"), oEmails("emailuserid"), sFromName, sFromEmail, sSubject, sMessageBody, oEmails("useremail")
     			iSentCount = iSentCount + 1
   		end if

   		oEmails.movenext
 	loop

 	oEmails.close
	 set oEmails = nothing

 	SendChildrenMail = iSentCount

end function

'------------------------------------------------------------------------------
sub subSendEmail( sToName, sEmailUserId, sFromName, sFromEmail, sSubject, sHTMLBody, sSendToEmail )

  if sSendToEmail <> "" then
	    lcl_from_email = sFromName & " <" & sFromEmail & ">"
  			'lcl_sendto     = sSendToEmail & " <" & sToName & ">"
  			lcl_sendto     = sSendToEmail
  			lcl_subject    = sSubject
  			lcl_html_body  = sHTMLBody

    'Send the email
     sendEmail lcl_from_email, lcl_sendto, "", lcl_subject, lcl_html_body, "", "Y"
 	end if

end sub

'------------------------------------------------------------------------------
sub dtb_debug(p_value)

  sSQL = "INSERT INTO my_table_dtb(notes) VALUES ('" & replace(p_value,"'","''") & "')"

 	set oDTB = Server.CreateObject("ADODB.Recordset")
	 oDTB.Open sSQL, Application("DSN"), 0, 1

end sub
%>
