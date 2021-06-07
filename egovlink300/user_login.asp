<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<% 
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: user_login.asp
' AUTHOR: ???
' CREATED: 2004
' COPYRIGHT: Copyright 2004 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Lets citizens log into the system.
'
' MODIFICATION HISTORY
' 1.0   2004	INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sError, oActionOrg
Dim iSectionID, sDocumentTitle, sURL, datDate, datDateTime, sVisitorIP

Set oActionOrg = New classOrganization


'Org Features
lcl_orghasfeature_show_actionline_links = orghasfeature(iorgid,"show actionline links")

'Check for org "edit displays"
lcl_orghasdisplay_donotknock_login_message = orghasdisplay(iorgid,"donotknock_login_message")

%>
<html>
<head>
  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<title>E-Gov Services <%=sOrgName%></title>

	<link rel="stylesheet" type="text/css" href="css/styles.css" />
	<link rel="stylesheet" type="text/css" href="global.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

	<script language="javascript" src="scripts/modules.js"></script>


</head>

<!--#Include file="include_top.asp"-->

<%

'This is for rerouting from another page that needs them logged in to function - Steve Loar 4/5/2006
lcl_loginDisplayMessage = ""
lcl_login_message       = ""

If session("LoginDisplayMsg") <> "" Then 
	lcl_loginDisplayMessage = "<p align=""center""><strong>" & session("LoginDisplayMsg") & "</strong></p>"
	session("LoginDisplayMsg") = ""
End If 

response.write vbcrlf & "<tr>"
response.write "<td valign=""top"">"
response.write vbcrlf & "<div class=""main"">"
response.write vbcrlf & "<font class=""pagetitle"">" & sWelcomeMessage & "</font><br />"
RegisteredUserDisplay( "" )

If lcl_orghasfeature_show_actionline_links Then 
	response.write "<a href=""action.asp"">"
	response.write "<img src=""images/arrow_2back.gif"" align=""absmiddle"" border=""0"" />"
	response.write "&nbsp;Return to " & oActionOrg.GetOrgFeatureName( "action line" )
	response.write "</a>"
	response.write "<br /><br />"
End If 

response.write vbcrlf & "</div>"
response.write vbcrlf & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" align=""center"">"
response.write vbcrlf & "<tr>"
response.write "<td align=""center"">"

response.write lcl_loginDisplayMessage

response.write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">"
'Determine if the user is wanting to sign up on the "Do Not Knock" list
If request("p") = "dnk" Then 
	If lcl_orghasdisplay_donotknock_login_message Then 
		lcl_login_message = getOrgDisplay(iorgid,"donotknock_login_message")
	Else 
		lcl_login_message = "<div style=""font-size:14px; font-weight:bold;"">DO NOT KNOCK REGISTRATION</div>"
		lcl_login_message = lcl_login_message & "<p style=""color:#800000"">Log into an existing, or register a new, account to<br />add yourself to the ""Do Not Knock"" list(s)</p>"
	End If 
End If 

'Show a message if one exists
If lcl_login_message <> "" Then 
	response.write vbcrlf & "<tr>"
	response.write "<td align=""center"">" & lcl_login_message & "</td>"
	response.write "</tr>"
End If 

response.write vbcrlf & "<tr>"
response.write "<td align=""center"">"
response.write vbcrlf & "<div class=""box_header2"">Sign In</div>"
response.write vbcrlf & "<div class=""groupSmall400"">"
%><!--#include file="inc_login.asp"--><%
response.write vbcrlf & "<tr>"
response.write "<td colspan=""2"" align=""left"">"
response.write "<p><br />"
response.write vbcrlf & "<ul>"
response.write vbcrlf & "<li><a href=""forgot_password.asp"">Can't remember your password?</a></li>"
response.write vbcrlf & "<li><a href=""register.asp"">Not registered yet?</a></li>"
response.write vbcrlf & "</ul>"
response.write "</p>"
response.write "</td>"
response.write "</tr>"
response.write vbcrlf & "<tr>"
response.write "<td colspan=""2"" align=""left"">" 
response.write vbcrlf & "<div id=""problemtextfield1"">"
response.write vbcrlf & "Internal Use Only, Leave Blank: <input type=""text"" name=""frmsubjecttext"" id=""problemtextinput"" value="""" size=""6"" maxlength=""6"" /><br />"
response.write vbcrlf & "<strong>Please leave this field blank and remove any <br />values that have been populated for it.</strong>"
response.write "</div>"
response.write "</td>"
response.write "</tr>"


response.write vbcrlf & "</table>"
response.write vbcrlf & "</form>"
response.write vbcrlf & "</div>"

response.write "</td>"
response.write "</tr>"

response.write vbcrlf & "</table>"

response.write "</td>"
response.write "</tr>"

response.write vbcrlf & "</table>"

response.write "</td>"
response.write "</tr>"

Set oActionOrg = Nothing 


'BEGIN: Visitor Tracking ------------------------------------------------------
iSectionID     = 1
sDocumentTitle = "MAIN"
sURL           = request.servervariables("SERVER_NAME") &":/" & request.servervariables("URL") & "?" & request.servervariables("QUERY_STRING")
datDate        = Date()	
datDateTime    = Now()
sVisitorIP     = request.servervariables("REMOTE_ADDR")

LogPageVisit iSectionID, sDocumentTitle, sURL, datDate, datDateTime, sVisitorIP, iorgid 
'END: Visitor Tracking -------------------------------------------------------

%>

<!--#Include file="include_bottom.asp"-->
