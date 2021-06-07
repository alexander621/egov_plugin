<!-- #include file="../includes/common.asp" //-->
<!--#Include file="class_global_functions.asp"-->  
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getuserinfo.asp
' AUTHOR: Steve Loar
' CREATED: 07/22/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets user information to be displayed on the class signup page. 
'				It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   07/22/2011   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iUserid, sSql, oRs, sResults, iCitizenUserId, bRegistrationBlocked, sBlockedValue

iCitizenUserId = CLng(0)

If request("citizenuserid") <> "" Then 
	If IsNumeric( request("citizenuserid") ) Then 
		iCitizenUserId = CLng("" & request("citizenuserid"))
	End If 
Else
	If session("eGovUserId") <> "" Then 
		If IsNumeric( Session("eGovUserId") ) Then 
			iCitizenUserId = Session("eGovUserId")
		End If 
	End If 
End If 
'iCitizenUserId = CLng(request("citizenuserid"))

sBlockedValue = "no"

sSql = "SELECT userfname, userlname, useraddress, useraddress2, usercity, userstate, "
sSql = sSql & " userzip, usercountry, useremail, userhomephone, ISNULL(residenttype,'R') AS residenttype, "
sSql = sSql & " userworkphone, userfax, userbusinessname, userpassword, "
sSql = sSql & " userregistered, residenttype, residencyverified, registrationblocked, "
sSql = sSql & " blockeddate, blockedadminid, blockedexternalnote, blockedinternalnote "
sSql = sSql & " FROM egov_users WHERE userid = " & iCitizenUserId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then 

	sResults = sResults & vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""signupuserinfo"">"
	sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"">Name:</td><td >" & oRs("userfname") & " " & oRs("userlname") 
	sResults = sResults & "&nbsp;&nbsp;&nbsp;<strong>" & GetResidentTypeDesc( oRs("residenttype") ) & "</strong>"

	If Not oRs("residencyverified") And oRs("residenttype") = "R" Then 
		If OrgHasFeature("residency verification") Then 
			sResults = sResults & " (not verified)"
		End If 
	End If 

	sResults = sResults & "</td></tr>"
	sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & oRs("useremail") & "</td></tr>"
	sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhone(oRs("userhomephone")) & "</td></tr>"
	sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"">Address:</td><td>" & oRs("useraddress") & "<br />" 

	If oRs("useraddress2") <> "" Then 
		sResults = sResults & oRs("useraddress2") & "<br />" 
	End If 

	If oRs("usercity") <> "" Or oRs("userstate") <> "" Or oRs("userzip") <> "" Then 
		  sResults = sResults & oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") 
	End If 

	sResults = sResults & "</td></tr>"

	' Handle blocked
	If oRs("registrationblocked") Then
		bRegistrationBlocked = True 
		sResults = sResults & vbcrlf & "<tr><td colspan=""2""><span id=""warningmsg""> *** Registration Blocked *** </span></td></tr>"
		sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"">Date:</td><td>" & oRs("blockeddate") & "</td></tr>"
		sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"">By:</td><td>" & GetAdminName( oRs("blockedadminid") ) & "</td></tr>"
		sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"" nowwrap=""nowrap"">Internal Note:</td><td>" & oRs.Fields("blockedinternalnote") & "</td></tr>"
		sResults = sResults & vbcrlf & "<tr><td align=""right"" valign=""top"" nowwrap=""nowrap"">External Note:</td><td>" & oRs.Fields("blockedexternalnote") & "</td></tr>"
	End If 

	sResults = sResults & vbcrlf & "</table>"

	If bRegistrationBlocked Then 
		sBlockedValue = "yes"
	End If 

Else
	sResults = sResults & vbcrlf & "<p>User information not found.</p>"
End If 

sResults = sResults & vbcrlf & "<input type=""hidden"" id=""registrationblocked"" name=""registrationblocked"" value=""" & sBlockedValue & """ />"

oRs.Close
Set oRs = Nothing

response.write sResults


%>