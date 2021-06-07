<%
intOrgID = 0
intUserID = 0
showGraph = false
Server.ScriptTimeout = 3600
LogThePage 0
if request.querystring("logid") <> "" then
	sSQL = "SELECT o.OrgName, o.OrgEgovWebsiteURL,u.UserID,l.*  " _
		& " FROM egov_class_distributionlist_log l  " _
		& " INNER JOIN Users u ON u.OrgID = l.orgid and u.IsRootAdmin = 1 and u.Username = 'eclink' " _
		& " INNER JOIN Organizations o ON o.OrgID = u.OrgID " _
		& " WHERE dl_logid = '" & dbsafe(request.querystring("logid")) & "'"
	Set oRs = Server.CreateObject("ADODB.RecordSet")
	oRs.Open sSQL, Application("DSN"), 3, 1
	if not oRs.EOF then
		'if now() > oRs("scheduledDateTime") then
			'for each item in oRs.fields
				'response.write item.name & " = " & item.value & "<br />"
			'next

			lcl_dl_logid = oRs("dl_logid")

			'CLEAR THE SCHEDULE SO IT DOESN'T RUN MORE THAN ONCE
    	    		clearSubscriptionInfoSchedule lcl_dl_logid

			response.cookies("User")("UserID") = oRs("UserID")
			response.cookies("User")("OrgID") = oRs("orgid")
			Session("orgid") = oRs("orgid")
			Session("userid") = oRs("userid")
			intOrgID = oRs("orgid")
			intUserID = oRs("userid")
			strOrgName = oRs("OrgName")
			strEGovURL = oRs("OrgEgovWebsiteURL")
	
	
    			subProcessEmail lcl_dl_logid
    			if request.form("sendlist") <> "" then SendPushNotification
    			updateSubscriptionInfoStatus lcl_dl_logid, "COMPLETED"
	


		'end if
	end if
	oRs.close
	Set oRs = Nothing
end if

Function DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	DBsafe = sNewString
End Function

'------------------------------------------------------------------------------------------------------------
' void LogThePage iLoadTime 
'------------------------------------------------------------------------------------------------------------
Sub LogThePage( ByVal iLoadTime )
	Dim sSql, oCmd, sScriptName, sVirtualDirectory, aVirtualDirectory, sPage, arr, sUserAgent, sUserAgentGroup

	sScriptName = Request.ServerVariables("SCRIPT_NAME")

	If request.servervariables("http_user_agent") <> "" Then 
		sUserAgent = "'" & DBsafe(Trim(Left(request.servervariables("http_user_agent"),480))) & "'"
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

	sSql = "INSERT INTO egov_pagelog ( virtualdirectory, applicationside, page, loadtime, scriptname, "
	sSql = sSql & " querystring, servername, remoteaddress, requestmethod, orgid, userid, username, useragent, useragentgroup,requestformcollection, cookiescollection, sessioncollection, sessionid ) VALUES ( "
	sSql = sSql & sVirtualDirectory & ", "
	sSql = sSql & "'admin', "
	sSql = sSql & "'" & sPage & "', "
	sSql = sSql & FormatNumber(iLoadTime,3) & ", "
	sSql = sSql & "'" & sScriptName & "', "

	If Request.ServerVariables("QUERY_STRING") <> "" Then 
		sSql = sSql & "'" & DBsafe(Left(Request.ServerVariables("QUERY_STRING"),500)) & "', "
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
	If session("orgid") <> "" Then 
		sSql = sSql & session("orgid") & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Userid
	If session("userid") <> "" Then
		sSql = sSql & session("userid") & ", "
	Else
		sSql = sSql & "NULL, "
	End If 

	' Get username
	If session("fullname") <> "" Then
		sSql = sSql & "'" & dbsafe(session("fullname")) & "', "
	Else
		sSql = sSql & "NULL, "
	End If 

	' User Agent
	sSql = sSql & sUserAgent & ", "

	' User Agent Group
	sSql = sSql & sUserAgentGroup & ", "
         'requestformcollection, cookiescollection, sessioncollection
	sSql = sSql & "'" & GetRequestFormCollection() & "',"
	sSql = sSql & "'" & GetCookiesCollection() & "',"
	sSql = sSql & "'" & GetSessionCollection() & "',"


	sSql = sSql & "'" & Session.SessionID & "'"

	sSql = sSql & " )"
	'response.write sSql
	session("PageLogSQL") = sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing
	session("PageLogSQL") = ""

End Sub 

Function GetRequestFormCollection()
	sPostLog = ""
	on error resume next
	For each item in Request.Form
		sPostLog = sPostLog & item & ":  " &	 request.form(item) & vbcrlf
	Next
	on error goto 0

	GetRequestFormCollection = dbsafe(sPostLog)
End Function
Function GetCookiesCollection()
	Collection = ""
	on error resume next
	For Each Item in Request.Cookies
		Collection = Collection & Item & ":  " & request.cookies(Item) & vbcrlf
	Next
	on error goto 0
	GetCookiesCollectionCollection = dbsafe(Collection)
End Function
Function GetSessionCollection()
	sSessionLog = ""
	on error resume next
	For each session_name in Session.Contents
		sSessionLog = sSessionLog & session_name & ":  " & session(session_name) & vbcrlf
	Next
	on error goto 0

	GetSessionCollection = dbsafe(sSessionLog)
End Function


'------------------------------------------------------------------------------------------------------------
' string GetUserAgentGroup( sUserAgent )
'------------------------------------------------------------------------------------------------------------
Function GetUserAgentGroup( ByVal sUserAgent )
	Dim sSql, oRs, sUserAgentGroup

	sUserAgentGroup = GetUntrackedUserAgentGroup()

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 0 ORDER BY checkorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		'If clng(InStr( sUserAgent, oRs("useragentgroup") )) > clng(0) Then
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


'------------------------------------------------------------------------------------------------------------
' string GetUntrackedUserAgentGroup()
'------------------------------------------------------------------------------------------------------------
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
'-------------------------------------------------------------------------------------------------
' void RunSQLStatement sSql 
'-------------------------------------------------------------------------------------------------
Sub RunSQLStatement( ByVal sSql )
	Dim oCmd

'	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush
	session("RunSQLStatement") = sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing
	
	session("RunSQLStatement") = ""

End Sub
'------------------------------------------------------------------------------
' Function UserIsRootAdmin( iUserID )
'------------------------------------------------------------------------------
Function UserIsRootAdmin( ByVal iUserID )
	Dim sSql, oRs

	UserIsRootAdmin = False 
	sSql = "SELECT ISNULL(isrootadmin,0) AS isrootadmin FROM users WHERE userid = " & iUserID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isrootadmin") Then 
			UserIsRootAdmin = True 
		End If 
	End If

	oRs.Close
	Set oRs = Nothing 

End Function 
'--------------------------------------------------------------------------------------------------
function BuildHTMLMessage(sBody,iContainsHTML)
	'Build email message
 	Dim sLayout, lcl_customhtml

 'Check to see if message contains HTML.
  if iContainsHTML <> "" then
     lcl_containsHTML = UCASE(iContainsHTML)
  else
     lcl_containsHTML = "N"
  end if

 'Check for custom html.  If any HTML tags exist in the email body the admin is sending then leave our HTML off of the email.
  lcl_customhtml = 0

  if instr(UCASE(sBody),"<HTML") > 0 then
     lcl_customhtml = lcl_customhtml + 1
  end if

  if instr(UCASE(sBody),"<HEAD") > 0 then
     lcl_customhtml = lcl_customhtml + 1
  end if

  if instr(UCASE(sBody),"<BODY") > 0 then
     lcl_customhtml = lcl_customhtml + 1
  end if

  if sBody <> "" then
    'If the message IS in HTML format already there is no need to set the carriage returns "chr(10)" to "<br />" tags.
    'If the message is NOT in HTML format then we want to make sure that the carriage returns are picked up 
     if lcl_containsHTML = "Y" then
        lcl_replace_linebreaks = ""
     else
        lcl_replace_linebreaks = "<br />"
     end if

     sBody = replace(sBody,vbcrlf,lcl_replace_linebreaks & vbcrlf)
  end if

  if lcl_customhtml < 1 then
    	lcl_return = lcl_return & "<html>" & vbcrlf
    	lcl_return = lcl_return & "<head>" & vbcrlf
    	'lcl_return = lcl_return & "  <STYLE> td {font-family: arial,tahoma; font-size: 12px; color: #000000;} </STYLE>"
    	lcl_return = lcl_return & "</head>" & vbcrlf
    	lcl_return = lcl_return & "<body bgcolor=""#efefef"">" & vbcrlf
    	lcl_return = lcl_return & "<font face=""helvetica, arial"">" & vbcrlf
    	lcl_return = lcl_return & "<p style=""margin:0px""></p>" & vbcrlf
    	lcl_return = lcl_return & "<table bordercolor=""#4A9E9F"" bgcolor=""#ffffff"" cellspacing=""0"" cellpadding=""5"" width=""95%"" align=""center"" border=""2"" valign=""top"">" & vbcrlf
     lcl_return = lcl_return & "<tr>" & vbcrlf
     lcl_return = lcl_return & "<td style=""font-family:arial,tahoma; font-size:12px; color:#000000;"">" & vbcrlf
  end if

  lcl_return = lcl_return & sBody & vbcrlf

  if lcl_customhtml < 1 then
    	lcl_return = lcl_return & "<center>" & vbcrlf
    	lcl_return = lcl_return & "<br />" & vbcrlf
    	lcl_return = lcl_return & "<hr color=""black"" size=""1"" width=""95%"">" & vbcrlf
    	lcl_return = lcl_return & "<font size=""-2"">Copyright 2004 - " & year(now) & ". <i>Electronic Commerce</i> Link, Inc. dba <i>EC</i> Link.</font>" & vbcrlf
     lcl_return = lcl_return & "</center>" & vbcrlf
     lcl_return = lcl_return & "</td>" & vbcrlf
     lcl_return = lcl_return & "</tr>" & vbcrlf
     lcl_return = lcl_return & "</table>" & vbcrlf
     lcl_return = lcl_return & "</font>" & vbcrlf
    	lcl_return = lcl_return & "</body>" & vbcrlf
    	lcl_return = lcl_return & "</html>" & vbcrlf
  end if

 	BuildHTMLMessage = lcl_return
end function

'------------------------------------------------------------------------------
' boolean OrgHasDisplay( iorgid, sDisplay )
'------------------------------------------------------------------------------
Function OrgHasDisplay( ByVal iOrgId, ByVal sDisplay )
	Dim sSql, oDisplay, blnReturnValue

	' SET DEFAULT
	blnReturnValue = False

	' LOOKUP passed display FOR the current ORGANIZATION 
	sSql = "SELECT COUNT(OD.displayid) AS display_count "
	sSql = sSql & " FROM egov_organizations_to_displays OD, egov_organization_displays D "
	sSql = sSql & " WHERE OD.displayid = D.displayid "
	sSql = sSql & " AND orgid = " & iOrgId
	sSql = sSql & " AND D.display = '" & sDisplay & "' "

	set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSql, Application("DSN"), 3, 1
	
	If clng(oDisplay("display_count")) > 0 Then
		' the ORGANIZATION HAS the Display
		blnReturnValue = True
	End If
	
	oDisplay.close 
	Set oDisplay = Nothing

	' set the RETURN  value
	OrgHasDisplay = blnReturnValue
End Function

'------------------------------------------------------------------------------
Function getOrganization_WP_URL( byVal iOrgId, byVal sColumnName )
  ' This pulls the Word Press URL for the passed org and column
  Dim sSql, oRs, sURL
  
  sURL = ""
  
  sSql = "SELECT " & sColumnName & " AS wp_url,OrgPublicWebsiteURL FROM organizations WHERE wpLive = 1 and orgid = " & CLng(iOrgId)
  'response.write sSQL & "<br /><br />"
  
  Set oRs = Server.CreateObject("ADODB.Recordset")
  oRs.Open sSql, Application("DSN"), 0, 1
  
  'response.write oRs("wp_url") & "<br /><br />"

  If Not oRs.EOF Then
    sURL = oRs("wp_url")
    if sURL = "" or isnull(sURL) then
    	if sColumnName = "wp_actionline_url" then
	    sURL = oRs("OrgPublicWebsiteURL") & "/citizen-action-line/"
    	elseif sColumnName = "wp_subscriptions_url" then
	    sURL = oRs("OrgPublicWebsiteURL") & "/subscriptions/"
    	end if
    end if
  End If
  
  oRs.Close
  Set oRs = Nothing  
  
  getOrganization_WP_URL = sURL
   
End Function


%>


<!--#Include file="inc_sendmail.asp"-->  
<!--#include file="../../egovlink300_global/includes/inc_email.asp" //-->

