<%
 response.write "      </td>" & vbcrlf
 response.write "      <td class=""respHide"" width=""1""><img src=""" & sImgBaseURL & "/img/clearshim.gif"" border=""0"" width=""1"" /></td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "</table>" & vbcrlf

'END: DIV for Midland ---------------------------------------------------------
 if iorgid = 29 then
    response.write "</div>" & vbcrlf
 end if

'END: Page Content ------------------------------------------------------------

	Dim oOrgb
	Set oOrgb = New classOrganization

'BEGIN: Fade Lines ------------------------------------------------------------
 response.write "<table id=""fadelines"" bgcolor=""#d6d3ce"" border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">" & vbcrlf
 response.write "  <tr bgcolor=""#666666""><td height=""1"" colspan=""2""></td></tr>" & vbcrlf
 response.write "  <tr bgcolor=""#ffffff""><td height=""1"" colspan=""2""></td></tr>" & vbcrlf
 response.write "</table>" & vbcrlf
'END: Fade Lines --------------------------------------------------------------

'BEGIN: Bottom menu and copyright informations --------------------------------
 response.write "<center>" & vbcrlf
 response.write "<div class=""footerbox"">" & vbcrlf
 response.write "<table width=""100%"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td valign=""top"" align=""center"">" & vbcrlf
 response.write "          <font class=""footermenu"">" & vbcrlf

 if blnFooterOn then
    response.write "          <br />" & vbcrlf

   'City Home (maintained in Org Properites)
    if oOrgb.checkMenuOptionEnabled("CITY") then
       lcl_label_city = oOrgb.getMenuOptionLabel("CITY")
       response.write "          <a href=""" & oOrgb.GetOrgURL() & """ class=""afooter"" target=""_top"">" & lcl_label_city & "</a> | " & vbcrlf
       'response.write "          <a href=""" & oOrgb.GetOrgURL()  & """ class=""afooter"" target=""_top"">" & oOrgb.GetOrgDisplayName("homewebsitetag") & "</a> |" & vbcrlf
    end if

   'E-Gov Home (maintained in Org Properties)
    if oOrgb.checkMenuOptionEnabled("EGOV") then
       lcl_label_egov = oOrgb.getMenuOptionLabel("EGOV")
       response.write "          <a href=""" & oOrgb.GetEgovURL() & """ class=""afooter"" target=""_top"">" & lcl_label_egov & "</a>" & vbcrlf
       'response.write "          <a href=""" & oOrgb.GetEgovURL() & """ class=""afooter"" target=""_top"">E-Gov Home</a>" & vbcrlf
    end if

                              oOrgb.ShowPublicFooterNav 2
    response.write "          <br />" & vbcrlf

    'if oOrgb.OrgHasDisplay( "privacy policy" ) then
    '   response.write "          <a href=""" & oOrgb.GetEgovURL() & "/privacy_policy_display.asp"" class=""afooter"" target=""_top""><strong>Privacy Policy</strong></a> |" & vbcrlf
    'end if

    if oOrgb.OrgHasDisplay( "refund policy" ) then
       response.write "          <a href=""" & oOrgb.GetEgovURL() & "/refund_policy.asp"" class=""afooter"" target=""_top"">Refund Policy</a> |" & vbcrlf
    end if

    response.write "          <a href=""" & oOrgb.GetEgovURL() & "/user_login.asp"" class=""afooter"" target=""_top"">Login</a> | <a href=""" & oOrgb.GetEgovURL() & "/register.asp"" class=""afooter"">Register</a>" & vbcrlf

    lcl_isDefaultPage     = false
    lcl_privacypolicy_url = displayPrivacyPolicyLink(lcl_isDefaultPage, iorgid)

    response.write lcl_privacypolicy_url & vbcrlf

 end if

'Copyright
 response.write "          <br />" & vbcrlf
 response.write "          <font class=""footermenu"">Copyright &copy;2004-" & year(now) & ". <em>Electronic Commerce</em> Link, Inc. " & vbcrlf
 response.write "          dba <a href=""http://www.egovlink.com"" target=""_NEW""><font class=""footermenu"">egovlink</font></a>.</font>" & vbcrlf

'BEGIN: Demo check to add admin link ------------------------------------------
 if oOrgb.OrgHasFeature( "AdministrationLink" ) then
    response.write "&nbsp;&nbsp;&nbsp;<a target=""_new"" href=""" & sEgovWebsiteURL & "/admin/"" class=""hidden"">Administrator</a>" & vbcrlf
 end if
'END: Demo check to add admin link --------------------------------------------

'BEGIN: Google Translator -----------------------------------------------------
If application("environment") = "PROD" Then 
	If oOrgb.orghasfeature("google_translator") Then 
		response.write "<div id=""google_translate_element""></div><script>" & vbcrlf
		response.write "function googleTranslateElementInit() {" & vbcrlf
		response.write "  new google.translate.TranslateElement({" & vbcrlf
		response.write "    pageLanguage: 'en'" & vbcrlf
		response.write "  }, 'google_translate_element');" & vbcrlf
		response.write "}" & vbcrlf 
		response.write "</script>"
		If request.servervariables("HTTPS") <> "on" Then
			response.write "<script src=""http://translate.google.com/translate_a/element.js?cb=googleTranslateElementInit""></script>" & vbcrlf
'		Else 
'			this is to handle the translator on our secure pages so no warning pops up about non-secure items on the page
'			response.write "<script src=""https://translate.google.com/translate_a/element.js?cb=googleTranslateElementInit""></script>" & vbcrlf
		End If 
	End If 
End If 
'END: Google Translator -------------------------------------------------------

 response.write "<br /><br />" & vbcrlf
 response.write "</font>" & vbcrlf
 response.write "</td></tr></table>" & vbcrlf
 response.write "</div>" & vbcrlf
 response.write "</center>" & vbcrlf
'END: Bottom menu and copyright information -----------------------------------

 set oOrgb = nothing 

 Dim iLoadTime
 iLoadTime = CDbl(0.00)

 if iStartSecs <> "" then
	   iLoadTime = timer - iStartSecs

   	if CDbl(iLoadTime) < CDbl(0.000) then
     		iLoadTime = CDbl(0.000)
   	end if
 end if

 LogThePage
 %>
 <script>
 function onElementHeightChange(elm, callback){
    var lastHeight = elm.clientHeight, newHeight;
    (function run(){
        newHeight = elm.clientHeight;
        if( lastHeight != newHeight )
            callback();
        lastHeight = newHeight;

        if( elm.onElementHeightChangeTimer )
            clearTimeout(elm.onElementHeightChangeTimer);

        elm.onElementHeightChangeTimer = setTimeout(run, 200);
    })();
}
	if (window.top!=window.self)
	{
		onElementHeightChange(document.body, function(){
			//alert("HERE");
			var height = document.body.scrollHeight;
 			parent.postMessage({event_id: 'heightchange',data: { heightval: height, initial: false }},"*")
			//console.log(height);
		});
		var height = document.body.scrollHeight;
		parent.postMessage({event_id: 'heightchange',data: { heightval: height, initial: true }},"*")
			//console.log(height);

	}
 </script>
 <%


 response.write "</body>" & vbcrlf
 response.write "</html>" & vbcrlf


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
	sSql = sSql & " servername, remoteaddress, requestmethod, orgid, userid, username, sectionid, documenttitle, useragent, useragentgroup, requestformcollection, cookiescollection, sessioncollection, sessionid  ) VALUES ( "
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
	If request.cookies("userid") <> "" and request.cookies("userid") <> "-1" and isnumeric(request.cookies("userid")) Then
		sSql = sSql & request.cookies("userid") & ", "
	Else
		sSql = sSql & "NULL, "
		response.cookies("userid") = ""
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

	sSql = sSql & "'" & Track_DBsafe(GetRequestformInformation()) & "',"
	sSql = sSql & "'" & GetCookiesCollection() & "',"
	sSql = sSql & "'" & GetSessionCollection() & "',"


	sSql = sSql & "'" & Session.SessionID & "'"

	sSql = sSql & " )"
	'response.write sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql

	session("sSql") = sSql
	oCmd.Execute
	session("sSql") = ""

	Set oCmd = Nothing


End Sub 
Function GetCookiesCollection()
	Collection = ""
	on error resume next
	For Each Item in Request.Cookies
		Collection = Collection & Item & ":  " & request.cookies(Item) & vbcrlf
	Next
	on error goto 0
	GetCookiesCollectionCollection = track_dbsafe(Collection)
End Function
Function GetSessionCollection()
	sSessionLog = ""
	on error resume next
	For each session_name in Session.Contents
		sSessionLog = sSessionLog & session_name & ":  " & session(session_name) & vbcrlf
	Next
	on error goto 0

	GetSessionCollection = track_dbsafe(sSessionLog)
End Function


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
