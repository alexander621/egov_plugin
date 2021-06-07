<!--#include file="../egovlink300_global/includes/inc_email.asp"-->
<%

code = request.querystring("code")


sSQL = "SELECT * FROM linktracker WHERE linkcode = '" & code & "'"
Set oRs = Server.CreateObject("ADODB.RecordSet")
oRs.Open sSQL, Application("DSN"), 3, 1
if not oRs.EOF then
	sSQL = "UPDATE linktracker SET clickeddate = '" & now() & "' WHERE linkcode = '" & code & "'"
	Set oCmd = Server.CreateObject("ADODB.Connection")
	oCmd.Open Application("DSN")
	oCmd.Execute(sSQL)
	oCmd.Close
	Set oCmd = Nothing

	'EMAIL MIKE
	sendEmail "sales@eclink.com","mgruber@eclink.com","","Link Tracking","""" & oRs("sendtoname") & """ clicked the """ & oRs("linktotext") & """ link you created on " & oRs("dategenerated"),"","N"

	

	response.redirect oRs("linktourl")
else
	response.write "Sorry, your link is broken."
end if
oRs.Close
Set oRs = Nothing
%>
