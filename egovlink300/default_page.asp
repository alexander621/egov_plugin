<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!-- #include file="includes/common.asp" //-->
<!-- #include file="includes/start_modules.asp" //-->
<!-- #include file="include_top_functions.asp" //-->
<!-- #include file="class/classOrganization.asp" //-->
<%
' Check if this is a mobile device and route accordingly
 'MobileCheck		' in include_top_functions.asp

'Determine if the org wants the "E-Gov Home" to be the CommunityLink.
'If so then redirect the user.
 lcl_confirm = "N"

 sSQL = "SELECT 'Y' AS lcl_confirm "
 sSQL = sSQL & " FROM egov_communitylink "
 sSQL = sSQL & " WHERE orgid = " & iorgid
 sSQL = sSQL & " AND isEgovHomePage = 1 "

 set oCLisHome = Server.CreateObject("ADODB.Recordset")
 oCLisHome.Open sSQL, Application("DSN"), 0, 1

 if not oCLisHome.eof then
    lcl_confirm = oCLisHome("lcl_confirm")
 end if

 oCLisHome.close
 set oCLisHome = nothing

 if lcl_confirm = "Y" then
    response.redirect "communitylink.asp"
 end if

 Dim sError, oOrg, iStartSecs

 iStartSecs = timer
 set oOrg   = New classOrganization

'Capture the current path
 session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME")

'Check for org features
 lcl_orghasfeature_google_translator  = orghasfeature(iorgid,"google_translator")
 lcl_orghasfeature_AdministrationLink = orghasfeature(iorgid,"AdministrationLink")

'Build the BODY "onload"
 lcl_onload = ""
%>
<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

	<title><%=oOrg.GetOrgName()%>, <%=oOrg.GetState()%> - Online Services</title>

	<link rel="stylesheet" type="text/css" href="css/default_page.css" />
	<link rel="stylesheet" type="text/css" href="css/style_<%=iorgid%>.css" />

<style type="text/css">

	div#banner {
		<%=replace(oOrg.GetOrgBannerBckgrnd(),"http://www.egovlink.com","")%>
		}

	@media screen and (max-width:480px)
	{
		#wrapper
		{
			background:none;
		}
		#nav
		{
			display:none;
		}
		#content
		{
			margin-left:0;
			width:100% !important;
			padding:0;
		}
		#banner
		{
			height:auto !important;;
		}
		#banner img
		{
			height:auto !important;
			width:100%;
		}
		div#wrapper
		{
			width:95%;
			margin-left:auto;
			margin-right:auto;
		}
		p.hasimage
		{
			min-height:53px;
			height:auto;
		}
</style>

<%
'Check to see if the org has the "google_translator" feature.  If so then build the required javascript code
'Also, add the function to the BODY "onload"
 if lcl_orghasfeature_google_translator then
    lcl_onload = lcl_onload & "googleTranslateElementInit();"
%>
<script src="https://translate.google.com/translate_a/element.js?cb=googleTranslateElementInit"></script>
<script>
function googleTranslateElementInit() {
  new google.translate.TranslateElement({
    pageLanguage: 'en'
  }, 'google_translate_element');
}
</script>
<% end if %>
<script>

	function iframecheck()
	{
 		if (window.top!=window.self)
		{
 			document.body.classList.add("iframeformat") // In a Frame or IFrame
 			//var element = document.getElementById("egovhead");
 			//element.classList.add("iframeformat");
		}
	}
</script>

</head>

<body onload="<%=lcl_onload & ";iframecheck();"%>">

<%
     	if request.servervariables("HTTPS") <> "on" then
	   sNavBaseURL = sEgovWebsiteURL
	else
	   sNavBaseURL = replace(sEgovWebsiteURL,"http://www.egovlink.com","https://secure.egovlink.com")
	end if
 response.write "<div id=""iframenav"" style=""display:none;"">"
 response.write "<div class=""iframenavlink iframenavbutton""><a href=""" & sNavBaseURL & "/rd_classes/class_categories.aspx"">Classes and Events</a></div>"
 response.write "<div class=""iframenavlink iframenavbutton""><a href=""" & sNavBaseURL & "/rentals/rentalcategories.asp"">Rentals</a></div>"
 response.write "<div class=""iframenavlink iframenavbutton""><a href=""" & sNavBaseURL & "/user_login.asp"">Login</a></div>"
 response.write "<div class=""searchMenuDiv"">    "
	 response.write "<div class=""searchBoxText iframenavbutton"" onClick=""expandiframeSearchBox()""><span>Search</span></div>    "
	 response.write "<div class=""searchBox"">      "
		 response.write "<div id=""iframeclassesSearchBox"" class=""classesSearchBox"" onmouseleave=""expandiframeSearchBox()"">"
			 response.write "<input type=""text"" id=""iframetxtsearchphrase"" name=""txtsearchphrase"" class=""txtsearchphrase"" value="""" size=""40"" />        "
			 response.write "<input type=""button"" name=""searchButton"" class=""searchButton"" value=""Find"" onClick=""iframeSearch()"" />      "
		 response.write "</div>    "
	 response.write "</div>  "
 response.write "</div>"
 response.write "</div>"
 response.write "<div id=""footerbug"" style=""display:none;""><a href=""http://www.egovlink.com"" target=""_top"">Powered By EGovLink</a></div>"
if iorgid = 37 or iorgid = 60 then 
	response.write "<script> function clearMsg(id) { if(document.getElementById('msg'+id)) { document.getElementById('msg'+id).style.display = ""none""; } } </script>"
	'response.write "<script type=""text/javascript"" src=""/eclink/rd_scripts/jquery-1.7.2.min.js""></script>"
	response.write "<script type=""text/javascript"" src=""/eclink/rd_scripts/egov_navigation_asp.js""></script>"
end if
 'Determine if there is a banner
  if oOrg.GetOrgBanner <> "" then
     response.write "<div id=""banner"">" & oOrg.GetOrgBanner() & "</div>" & vbcrlf
  end if

  response.write "<div id=""wrapper"">" & vbcrlf
  response.write "<div id=""nav"">" & vbcrlf
  response.write "  <div id=""subnav"">" & vbcrlf

  if oOrg.checkMenuOptionEnabled("CITY") then
     lcl_label_city = oOrg.getMenuOptionLabel("CITY")

     'response.write "<li><a href=""" & oOrg.GetOrgURL() & """>" & oOrg.GetOrgDisplayName("homewebsitetag") & "</a></li>" & vbcrlf
     response.write "<p><a href=""" & oOrg.GetOrgURL() & """>" & lcl_label_city & "</a></p>" & vbcrlf
  end if

  if oOrg.checkMenuOptionEnabled("EGOV") then
     lcl_label_egov = oOrg.getMenuOptionLabel("EGOV")
     response.write "<p><a href=""" & oOrg.GetEgovURL() & """>" & lcl_label_egov & "</a></p>" & vbcrlf
  end if
%>
				<!--<p><a href="<%=oOrg.GetOrgURL()%>"><%=oOrg.GetOrgDisplayName("homewebsitetag")%></a></p>-->
				<!--<p><a href="<%=oOrg.GetEgovURL()%>">E-Gov Home</a></p>-->
				<% oOrg.ShowPublicLeftNav %>
			</div>
			<div class="spacer">&nbsp;</div>

<%
 'BEGIN: Google Translator ----------------------------------------------------
  if lcl_orghasfeature_google_translator then
     response.write "   <center>" & vbcrlf
     response.write "     <span id=""google_translate_element"" style=""background-color:#ffffff; padding:2px 2px""></span>" & vbcrlf
     response.write "   </center>" & vbcrlf
  end if
 'END: Google Translator ------------------------------------------------------
%>
		</div>

		<div id="content">
			<div class="welcome">Welcome to the <%=oOrg.GetOrgName()%><% If (oOrg.GetState() <> "") Then
					response.write ", " & oOrg.GetState() & ","
				End If 
			
			%> e-Government Services Website.
			
			</div>
			<div class="datetagline">Today is <%=FormatDateTime(Date(), vbLongDate)%>. <%=sTagline%> </div>

<%

			If sOrgRegistration And request.cookies("userid") <> "" and request.cookies("userid") <> "-1" Then
				response.write vbcrlf & "<div id=""accountmenu"">"
				ShowLoggedinLinks ""
				response.write vbcrlf & "</div>"
			End If 

			If oOrg.OrgHasDisplay( "homepage message" ) Then 
				response.write vbcrlf & "<div id=""homepagemessage"">"
				response.write vbcrlf & oOrg.GetOrgDisplay( "homepage message" )
				response.write vbcrlf & "</div>"
			End If 

%>
			<% oOrg.ShowPublicMainNav %>

			<div class="spacer">&nbsp;</div>
		</div>

		<div id="footer">
			<p>
<%
 'City Home (maintained in Org Properites)
  if oOrg.checkMenuOptionEnabled("CITY") then
     lcl_label_city = oOrg.getMenuOptionLabel("CITY")
     response.write "<a href=""" & oOrg.GetOrgURL() & """ class=""adefaultfooter"">" & lcl_label_city & "</a> | " & vbcrlf
		   'response.write "<a href=""" & oOrg.GetOrgURL() & """ class=""adefaultfooter"">" & oOrg.GetOrgDisplayName("homewebsitetag") & "</a> |" & vbcrlf
  end if

 'E-Gov Home (maintained in Org Properties)
  if oOrg.checkMenuOptionEnabled("EGOV") then
     lcl_label_egov = oOrg.getMenuOptionLabel("EGOV")
     response.write "<a href=""" & oOrg.GetEgovURL() & """ class=""adefaultfooter"">" & lcl_label_egov & "</a>" & vbcrlf
     'response.write "<a href=""" & oOrg.GetEgovURL() & """ class=""adefaultfooter"">E-Gov Home</a>" & vbcrlf
  end if

  oOrg.ShowPublicDefaultFooterNav 2

  response.write "</p>" & vbcrlf
  response.write "<p>" & vbcrlf

  'if oOrg.OrgHasDisplay("privacy policy") then
  '   response.write "<a href=""" & oOrg.GetEgovURL() & "/privacy_policy_display.asp"" class=""adefaultfooter""><strong>Privacy Policy</strong></a> | " & vbcrlf
  'end if

  if oOrg.OrgHasDisplay("refund policy") then
     response.write "<a href=""" & oOrg.GetEgovURL() & "/refund_policy.asp"" class=""adefaultfooter"">Refund Policy</a> | " & vbcrlf
  end if

  lcl_isDefaultPage     = true
  lcl_privacypolicy_url = displayPrivacyPolicyLink(lcl_isDefaultPage, iorgid)

  response.write "<a href=""user_login.asp"" class=""adefaultfooter"">Login</a>" & vbcrlf
  response.write " | " & vbcrlf
  response.write "<a href=""register.asp"" class=""adefaultfooter"">Register</a>" & vbcrlf
  response.write lcl_privacypolicy_url & vbcrlf
  response.write "</p>" & vbcrlf

 'Link around egovlink added back per Peters request 12/12/2007
  response.write "<p>Copyright &copy; 2004-" & year(now) & " Electronic Commerce Link, Inc. dba <a href=""http://www.egovlink.com"" target=""_NEW"" class=""adefaultfooter"">E-Gov Link</a>" & vbcrlf

 'BEGIN: Demo Check ro add admin link ---------------------------------------
  if lcl_orghasfeature_AdministrationLink then
     response.write "&nbsp;&nbsp;&nbsp;<a target=""_new"" href=""" & sEgovWebsiteURL & "/admin/"" class=""hidden"">Administrator</a>" & vbcrlf
  end if
 'END: Demo Check ro add admin link -----------------------------------------

  response.write "</p>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</div>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

  set oOrg = nothing

  Dim iLoadTime
  iLoadTime = CDbl(0.00)

  if iStartSecs <> "" then
    	iLoadTime = timer - iStartSecs
    	'response.write FormatNumber(iLoadTime,3) & " seconds"
  end if

  LogThePage

'------------------------------------------------------------------------------
Sub LogThePage()
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

	sSql = "INSERT INTO egov_pagelog ( virtualdirectory, applicationside, page, loadtime, scriptname, "
	sSql = sSql & " querystring, servername, remoteaddress, requestmethod, orgid, userid, username, useragent, useragentgroup ) VALUES ( "
	sSql = sSql & sVirtualDirectory & ", "
	sSql = sSql & "'public', "
	sSql = sSql & "'" & sPage & "', "
	sSql = sSql & FormatNumber(iLoadTime,3) & ", "
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

	' User Agent
	sSql = sSql & sUserAgent & ", "

	' User Agent Group
	sSql = sSql & sUserAgentGroup


	sSql = sSql & " )"
	'response.write sSql

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub


'------------------------------------------------------------------------------
Function GetUserAgentGroup( ByVal sUserAgent )
	Dim sSql, oRs, sUserAgentGroup

	sUserAgentGroup = GetUntrackedUserAgentGroup()

	sSql = "SELECT useragentgroup FROM UserAgent_Groups WHERE isuntracked = 0 ORDER BY checkorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If clng(InStr( sUserAgent, oRs("useragentgroup") )) > clng(0) Then
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



%>
