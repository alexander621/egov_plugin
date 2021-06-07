<!-- #include file="adovbs.inc" -->
<!-- #include file="../../egovlink300_global/includes/inc_upload.asp" //-->
<!-- #include file="../../egovlink300_global/includes/inc_email.asp" //-->

<!-- #include file="../class/classOrganization.asp" -->
<%
' to turn the system off and redirect to the outage.html page, set the 0 to 1 below
SystemOutage 0

'LEAVE THIS COMMENTED OUT, Chip - 4/17/02    'Option Explicit
Response.Buffer = True


Dim sTmp, oAutoLoginCmd, sTimeOutPath, sLevel, iPageLogStartSecs
iPageLogStartSecs = timer 
'On Error Resume Next
RootPath = request.servervariables("URL")
RootPath = LEFT(RootPath,INSTR(2,RootPath,"/")-1)
RootPath = RootPath & "/admin/"
'response.write "<!--" & ROOTPath& "-->"

sLevel = ""  'Each page needs to override this to get the header and menu to work right

'Some constants used on almost every page
'Const adExecuteNoRecords = 128
'Const adCmdStoredProc = 4
'Const adCmdText = 1
'Const adInteger = 3
'Const adVarChar = 200
'Const adLongVarChar = 201
Const adDateTime = 135
'Const adParamReturnValue = 4
'Const adParamInput = 1
'Const adParamOutput = 2
'Const adOpenStatic = 3
'Const adUseClient = 3
'Const adLockReadOnly = 1
'Const adStateOpen = 1

Const tabCount = 14
Const tabHome =			1
Const tabCalendar  =	2
Const tabMessages =		3
Const tabDocuments =	4
Const tabCommittees =	5
Const tabDiscussions =	6
Const tabVoting =		7
Const tabMeetings =		8
Const tabAdmin =		9
Const tabActionline =   10
Const tabPayments	=   11
Const tabRequests =		12
Const tabRegistration = 13
Const tabRecreation =	14

Const ITEM_TYPE_DOCUMENT    = 1
Const ITEM_TYPE_VOTE        = 2
Const ITEM_TYPE_DISCUSSION  = 3
Const ITEM_TYPE_TEXT        = 4

Const MAILBOX_IN            = 1
Const MAILBOX_DRAFT         = 2
Const MAILBOX_SENT          = 3

Const COMPOSE_TYPE_REPLY    = 1
Const COMPOSE_TYPE_REPLYALL = 2
Const COMPOSE_TYPE_FOWARD   = 3

'If not logged in, check if user has stored cookie
sTmp = Request.Cookies("User")("UserID") & ""

'Check to see if user is logged in already
If Session("UserID") = 0 Or Session("UserID") = "" Then

  If sTmp <> "" and AuditCookie(sTmp) Then  'If they have a valid cookie
    Session("UserID") = CLng(sTmp)
    Session("OrgID") = Request.Cookies("User")("OrgID")
    Session("FullName") = Request.Cookies("User")("FullName") & ""
    'Session("PageSize") = Request.Cookies("User")("PageSize")
	Session("LocationId") = Request.Cookies("User")("LocationId")
    Session("ShowStockTicker") = Request.Cookies("User")("ShowStockTicker")
    Session("Permissions") = Request.Cookies("User")("Permissions") & ""


	Dim iorgid,iPaymentGatewayID,blnOrgRegistration,blnQuerytool,blnFaq
	Dim sorgVirtualSiteName, sOrgFormLetterOn,sOrgInternalEntry
	Dim iTimeOffset,sEgovWebsiteURL,blnSeparateIndex

	SetOrganizationParameters()

	' ORG PARMS
	session("payment_gateway") = iPaymentGatewayID
	session("orgregistration") = blnOrgRegistration
	session("orgquerytool") = blnQueryTool
	session("orgfaq") = blnFaq
	session("sitename") = sorgVirtualSiteName
	session("OrgFormLetterOn") = sOrgFormLetterOn
	session("OrgInternalEntry") = sOrgInternalEntry
	session("iTimeOffset") = iTimeOffset
	session("egovclientwebsiteurl") = sEgovWebsiteURL
	session("virtualdirectory") = sorgVirtualSiteName
	

    Set oAutoLoginCmd = Server.CreateObject("ADODB.Command")
    With oAutoLoginCmd
      .ActiveConnection = Application("DSN")
      .CommandText = "NewAuditEvent"
      .CommandType = adCmdStoredProc
      .Parameters.Append .CreateParameter("UserID", adInteger, adParamInput, 4, sTmp)
      .Parameters.Append .CreateParameter("Type", adVarChar, adParamInput, 25, "Login (Auto)")
      .Parameters.Append .CreateParameter("Object", adVarChar, adParamInput, 50, Session.SessionID)
      .Parameters.Append .CreateParameter("Notes", adVarChar, adParamInput, 200, Request.ServerVariables("REMOTE_ADDR"))
      .Execute
    End With
    Set oAutoLoginCmd = Nothing
  Else
    'If no stored cookie then send to login page if required
    If Application("LogInRequired") = True Then
      If PageIsRequiredByLogin = False Then
        'keep initial page request in memory
        Session("RedirectPage") = Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString()
        'redirect to login page
        Response.Redirect RootPath & "login.asp"
      End If
    End If
  End If
End If

'Kick Out Anybody that isn't enabled
twfURL = request.servervariables("URL")
if instr(twfURL,"signoff.asp") = 0 and request.cookies("User")("UserID") <> "" then
	sSQL = "SELECT enabled FROM Users WHERE enabled = 1 and isdeleted = 0 AND userid = '" & replace(request.cookies("User")("UserID"),"'","''") & "'"
	set oCheck = Server.CreateObject("ADODB.RecordSet")
	oCheck.Open sSQL, Application("DSN"), 3, 1
	if oCheck.EOF then 
		response.redirect RootPath & "signoff.asp"
	end if
	oCheck.Close
	Set oCheck = Nothing
end if

if Session("OrgID") = "153" then
	'response.write Request.ServerVariables("REMOTE_ADDR")
	    'AND Request.ServerVariables("REMOTE_ADDR") <> "" _
	if Request.ServerVariables("REMOTE_ADDR") <> "74.87.250.138" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "72.49.14.92" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "76.190.86.58" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "96.250.72.98" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "75.99.210.50" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "108.58.203 130" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "71.183.66.210" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "75.127.145.123" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "96.57.250.162" _
	    AND Request.ServerVariables("REMOTE_ADDR") <> "173.220.236.234" _
		Then
		'Response.Redirect "http://www.ryeny.gov"
	end if
end if



%>
<!-- #include file="../custom/includes/custom.asp" //-->
<%
'If User is not a guest and they dont log in automatically, write a script to prevent Session Timeout
' DISABLED
If Session("UserID") > 0 And sTmp = "" AND 1=2 Then
%>
  <script language="Javascript">
    function mySetTimeout() {
      window.setTimeout("onTimeout()", 60000*15);  //15 minutes til popup, then 5 more til logout if no response
    }
    function onTimeout() {
      window.open("<%=RootPath%>timeoutPopup.asp", "_timeout", "alwaysRaised=yes,height=100,width=400,location=no,menubar=no,resizeable=no,scrollbars=no,resizeable=no,status=no,left=" + (screen.width-400)/2 + ",top=" + (screen.height-100)/2)
    }
    window.onload = mySetTimeout;
  </script>
<%
End If

'----------------------------------------------------------------------------------------
'	void SystemOutage iOutageFlag
'----------------------------------------------------------------------------------------
Sub SystemOutage( ByVal iOutageFlag )
	' this will redirect to the outage page so entire system can be shut down.
	Dim sRedirect

	If clng(iOutageFlag) = clng(1) Then
		If session("UserID") <> "" Then 
			sUserId = CLng(session("UserID"))
		Else
			sUserId = CLng(0) 
		End If 

		' see if they are the root admin and if so let them continue. For eclink they are dev=1408, prod=1710, test=1710
		If sUserId <> CLng("1710") Then 
			sRedirect = "http://" + request.ServerVariables("server_name") + "/" + GetVirtualDirectyName() + "/admin/outage.html"
			response.redirect sRedirect
		End If 
	End If 

End Sub 

'----------------------------------------------------------------------------------------
' boolean HasPermission( PrivilegeName )
'----------------------------------------------------------------------------------------
'Check to see if user has permission to access a specific function
Function HasPermission( ByVal PrivilegeName )

	If InStr(1, Session("Permissions"), PrivilegeName) Then
		HasPermission = True
	Else
		HasPermission = False
	End If

End Function

'Draw Standard Tab Heading (Which tab to put in front is passed as a param)
'Also the directory level is passed, so we know how many "../" -s we need


'------------------------------------------------------------------------------
' boolean OrgHasFeature( sFeature )
'------------------------------------------------------------------------------
Function OrgHasFeature( ByVal sFeature )
	Dim sSql, oRs

	' LOOKUP passed FEATURE FOR the current ORGANIZATION 
	sSql = "SELECT FO.featureid "
	sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid AND FO.orgid = " & Session("OrgID")
	sSql = sSql & " AND F.feature = '" & sFeature & "' "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		OrgHasFeature = True
	Else
		OrgHasFeature = False
	End If 
	
	oRs.Close 
	Set oRs = Nothing

End Function


'------------------------------------------------------------------------------
' string GetFeatureURL( sFeature )
'------------------------------------------------------------------------------
Function GetFeatureURL( ByVal sFeature )
	Dim sSql, oFeature

	' LOOKUP the URL FOR feature SUPPLIED
	sSql = "SELECT adminurl FROM egov_organization_features WHERE feature = '" & sFeature & "' "

	Set oFeature = Server.CreateObject("ADODB.Recordset")
	oFeature.Open  sSql, Application("DSN"), 0, 1
	
	If Not oFeature.EOF Then
		GetFeatureURL = oFeature("adminurl")
	Else
		GetFeatureURL = "#"
	End If
	
	oFeature.Close 
	Set oFeature = Nothing

End Function


'----------------------------------------------------------------------------------------
' string GetFeatureName( sFeature )
'----------------------------------------------------------------------------------------
Function GetFeatureName( ByVal sFeature )
	Dim sSql, oFeature

	sSql = "SELECT ISNULL(FO.featurename,F.featurename) AS featurename "
	sSql = sSql & " FROM egov_organizations_to_features FO, egov_organization_features F "
	sSql = sSql & " WHERE FO.featureid = F.featureid AND FO.orgid = " & session("OrgId")
	sSql = sSql & " AND F.feature = '" & sFeature & "'" 

	Set oFeature = Server.CreateObject("ADODB.Recordset")
	oFeature.Open sSql, Application("DSN"), 3, 1

	If Not oFeature.EOF Then
		GetFeatureName = oFeature("featurename")
	Else
		GetFeatureName = ""
	End If 

	oFeature.Close
	Set oFeature = Nothing 

End Function 


'----------------------------------------------------------------------------------------
' string GetGoogleApiKey()
'----------------------------------------------------------------------------------------
Function GetGoogleMapApiKey()
	Dim sSql, oRs

	sSql = "SELECT googlemapapikey FROM organizations WHERE orgid = " & session("OrgId")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetGoogleMapApiKey = oRs("googlemapapikey")
	Else
		GetGoogleMapApiKey = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' Sub DrawTabs( FrontTab, dirLevel )
'------------------------------------------------------------------------------
Sub DrawTabs( ByVal FrontTab, ByVal dirLevel )

	Dim sMsg, i, sLink, sWords, bPrevInFront, sRelDir, sSql, oNav, iFrontTab

	'sRelDir=left("../../../../../../",dirLevel*3)
	sRelDir = RootPath

	sMsg = "        <table border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & vbCRLF
	' Display customer logo
	sMsg = sMsg & "<td valign=top background=""" & sRelDir & "images/logo_background.jpg"" ><img src='" & sRelDir & custGraphic & "home.jpg'></td>"

	If Session("UserID") <> 0 Then 
		' Get the available admin tabs for this organization
		sSql = "Select O.OrgEgovWebsiteURL, F.adminURL, F.featurename, isnull(F.fronttabno,0) as fronttabno "
		sSql = sSql & "FROM organizations O, egov_organizations_to_features FO, egov_organization_features F "
		sSql = sSql & "WHERE (FO.hasadminview = 1 or F.hasadminview = 1) and parentfeatureid = 0 "
		sSql = sSql & " and O.orgid = FO.orgid and FO.featureid = F.featureid and O.orgid = " & Session("OrgID")
		sSql = sSql & " Order By FO.admindisplayorder, F.admindisplayorder"

		Set oNav = Server.CreateObject("ADODB.Recordset")
		oNav.Open sSql, Application("DSN"), 3, 1

		Do While Not oNav.EOF 
			sMsg = sMsg & "<td><img src=""" & sRelDir
			'sLink = oNav("OrgEgovWebsiteURL") & "/admin/" & oNav("adminURL")
			sLink =  oNav("adminURL")
			sWords = oNav("featurename")
			iFrontTab = clng(oNav("fronttabno")) 

			' leading tab edge
			If iFrontTab = 1 Then   'Home tab gets special "home" handling:  Graphic is slightly different
				sMsg = sMsg & "images/tabfront"
				If FrontTab = 1 Then 
					sMsg = sMsg & "_selected"
					bPrevInFront = True
				End If 
			Else 
				' tab edges for the middle tabs
				If bPrevInFront Then 
					sMsg = sMsg & "images/tabright_selected"
					bPrevInFront = False
				Else 
					sMsg = sMsg & "images/lefttab"
					If FrontTab = iFrontTab Then 
						sMsg = sMsg & "_selected"
						bPrevInFront = True
					End If 
				End If 	 
			End If 
			sMsg = sMsg & ".jpg""></td><td background=""" & sRelDir & "images/back"
			If FrontTab = iFrontTab Then
				sMsg = sMsg & "_selected"
			End If 
			sMsg = sMsg & ".jpg""><a href=""" & sRelDir & sLink & """ class=""tab"" target=""_top"">" & sWords & "</a></td>" & vbCRLF
			oNav.MoveNext
		Loop 
		oNav.close
		Set onav = Nothing 
   End If   'end Skip tab logic

  Response.Write "  <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""menu"">"
  Response.Write "    <tr>"
  Response.Write "      <td height=""15""></td>"
  Response.Write "    </tr>"
  Response.Write "    <tr>"
  Response.Write "      <td colspan=""2"" background=""" & sRelDir  & "images/back_main.jpg"">"
  Response.Write sMsg
  Response.Write "            <td>"

  ' Ending tab edge
  If Session("UserID") <> 0 Then 
	  response.write "<img src=""" & sRelDir & "images/tabend"
	  if bPrevInFront  then
		Response.Write "_selected"
	  end if
	  Response.Write ".jpg"">"
  Else
	response.write "&nbsp;"
  End If 
  response.write "</td>"
  if session("UserID")=0 or session("UserID")=""  then
    Response.Write "            <td id=""Sign Off"">&nbsp;&nbsp;<a href=""" & sRelDir & "login.asp"" class=""tab"" style=""color:#ffffff;"" target=""_top"">" & langLogIn & "</a></td>"
  else
    Response.Write "            <td id=""Sign Off"">&nbsp;&nbsp;<a href=""" & sRelDir & "signoff.asp"" class=""tab"" style=""color:#ffffff;"" target=""_top"">" & langLogOut & "</a></td>"
  end if
  Response.Write "          </tr>"
  Response.Write "        </table>"
  Response.Write "      </td>"
  Response.Write "    </tr>"
  Response.Write "    <tr>"

  if FrontTab=tabDocuments then
    Response.Write "      <td class=""submenu""><img src=""images/spacer.gif"" width=""8"" height=""1""></td>"
    Response.Write "      <td class=""submenu"" width=""100%"" style=""padding-top:2px;"">"
    Response.Write "        <span class=""button"" onclick=""toggleMenu();"" onmouseover=""this.className='buttona';status='" & langCollapseMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/collapse.gif"" id=""img_toggle"" width=""18"" height=""18"" border=""0"" alt=""" & langCollapseMenu & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraToc.document.location.reload()"" onmouseover=""this.className='buttona';status='" & langRefreshMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/refresh.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langRefreshMenu & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=main.asp'"" onmouseover=""this.className='buttona';status='Home';"" onmouseout=""this.className='button';status='';""><img src=""images/document_home.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langDocumentsHome & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraTopic.focus();parent.fraTopic.print();"" onmouseover=""this.className='buttona';status='" & langPrintCurrDoc & "';"" onmouseout=""this.className='button';status='';""><img src=""images/print.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langPrintCurrDoc & """></span>"
    Response.Write "        &nbsp;"

    If HasPermission("CanEditDocuments") Then
     ' Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addfolder.asp'"" onmouseover=""this.className='buttona';status='" & langAddFolder & "';"" onmouseout=""this.className='button';status='';""><img src=""images/folder_closed.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddFolder & """></span>"
     ' Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addarticle.asp'"" onmouseover=""this.className='buttona';status='" & langAddDocument & "';"" onmouseout=""this.className='button';status='';""><img src=""images/document.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddDoc & """></span>"
    '  Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addhelp.asp'"" onmouseover=""this.className='buttona';status='" & langAddHelp & "';"" onmouseout=""this.className='button';status='';""><img src=""images/helpdocument.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddHelp & """></span>"
       Response.Write "          <span class=""button"" onclick=""parent.fraToc.ToggleContextMenu();"" onmouseover=""this.className='buttona';status='" & langAddToggle & "';"" onmouseout=""this.className='button';status='';""><img src=""images/menu.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddToggle & """></span>"
      Response.Write "          &nbsp;"
      'JS 3/28/02 Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=admin.asp'"" onmouseover=""this.className='buttona';status='" & langAdmin & "';"" onmouseout=""this.className='button';status='';""><img src=""images/admin.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAdmin & """></span>"
    End If
    Response.Write "      </td>"
  else
    Response.Write "      <td class=""submenu""><img src=""" & sRelDir & "images/spacer.gif"" height=""28"" width=""1"" border=""0""></td>"
  end if

  Response.Write "    </tr>"
  Response.Write "  </table>"
  Response.Write sTimeOutScript
End Sub


'------------------------------------------------------------------------------
' Sub oldDrawTabs( FrontTab, dirLevel )
'------------------------------------------------------------------------------
Sub oldDrawTabs( FrontTab, dirLevel )

  dim sMsg, i, sLink, sWords, bPrevInFront, sRelDir

  'sRelDir=left("../../../../../../",dirLevel*3)
  sRelDir = RootPath

  sMsg = "        <table border=""0"" cellpadding=""0"" cellspacing=""0""><tr>" & vbCRLF
  ' Display customer logo
  sMsg = sMsg & "<td valign=top background=""" & sRelDir & "images/logo_background.jpg"" ><img src='" & sRelDir & custGraphic & "home.jpg'></td>"

  ' Begin iterating thru the various tabs
  for i = 1 to tabCount
    if mid(custTabVisible,i,1)="Y" then  'Use logic like this to implement FGA, and skip sections

      if NOT(session("UserID")=0 and (i=tabMessages Or i=tabAdmin)) then 'dont draw the messages tab if not logged in

        sMsg = sMsg & "<td><img src=""" & sRelDir

        select case i
          case tabHome:
            sLink = ""
            sWords = "Home"
		  case tabCalendar:
            sLink = "events/"
            sWords = langTabHome
          case tabMessages:
            sLink = "messages/"
            sWords = langTabMessages
          case tabDocuments:
            sLink = "docs/"
            sWords = langTabDocuments
          case tabCommittees:
            sLink = "dirs/"
            sWords = langTabCommittees
          case tabDiscussions:
            sLink = "discussions/"
            sWords = langTabDiscussions
          case tabVoting:
            sLink = "polls/"
            sWords = langTabVoting
          case tabMeetings:
            sLink = "meetings/"
            sWords = langTabMeetings
          case tabAdmin:
            sLink = "admin/"
            sWords = langTabAdmin
		  case tabActionline:
            sLink = "action_line/"
            sWords = "Action Line"
		  case tabPayments:
            sLink = "payments/"
            sWords = "Payments"
		  case tabRequests:
			sLink = "action_line/action.asp"
			sWords = "New Request"
		  case tabRegistration
		  	sLink = "dirs/display_citizen_groups.asp"
			sWords = "Registration"
		  case tabRecreation
			sLink = "recreation/default.asp"
			sWords = "Recreation"

		end select

        if i=1 then  'First tab gets special "home" handling:  Graphic is slightly different
          sMsg = sMsg & "images/tabfront"
          if FrontTab = 1 then
            sMsg = sMsg & "_selected"
            bPrevInFront = True
          end if
        else
          if bPrevInFront then
            sMsg = sMsg & "images/tabright_selected"
            bPrevInFront = False
          else
            sMsg = sMsg & "images/lefttab"
            if FrontTab = i then
              sMsg = sMsg & "_selected"
              bPrevInFront = True
            end if
          end if
        end if

        sMsg = sMsg & ".jpg""></td><td background=""" & sRelDir & "images/back"
        if FrontTab = i then sMsg = sMsg & "_selected"
        sMsg = sMsg & ".jpg""><a href=""" & sRelDir & sLink & """ class=""tab"" target=""_top"">" & sWords & "</a></td>" & vbCRLF

      end if
   end if  'end Skip tab logic
  next

  Response.Write "  <table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""100%"" class=""menu"">"
  Response.Write "    <tr>"
  Response.Write "      <td height=""15""></td>"
  Response.Write "    </tr>"
  Response.Write "    <tr>"
  Response.Write "      <td colspan=""2"" background=""" & sRelDir  & "images/back_main.jpg"">"
  Response.Write sMsg
  Response.Write "            <td><img src=""" & sRelDir & "images/tabend"

  if bPrevInFront  then
    Response.Write "_selected"
  end if

  Response.Write ".jpg""></td>"
  if session("UserID")=0 or session("UserID")=""  then
    Response.Write "            <td id=""Sign Off"">&nbsp;&nbsp;<a href=""" & sRelDir & "login.asp"" class=""tab"" style=""color:#ffffff;"" target=""_top"">" & langLogIn & "</a></td>"
  else
    Response.Write "            <td id=""Sign Off"">&nbsp;&nbsp;<a href=""" & sRelDir & "signoff.asp"" class=""tab"" style=""color:#ffffff;"" target=""_top"">" & langLogOut & "</a></td>"
  end if
  Response.Write "          </tr>"
  Response.Write "        </table>"
  Response.Write "      </td>"
  Response.Write "    </tr>"
  Response.Write "    <tr>"

  if FrontTab=tabDocuments then
    Response.Write "      <td class=""submenu""><img src=""images/spacer.gif"" width=""8"" height=""1""></td>"
    Response.Write "      <td class=""submenu"" width=""100%"" style=""padding-top:2px;"">"
    Response.Write "        <span class=""button"" onclick=""toggleMenu();"" onmouseover=""this.className='buttona';status='" & langCollapseMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/collapse.gif"" id=""img_toggle"" width=""18"" height=""18"" border=""0"" alt=""" & langCollapseMenu & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraToc.document.location.reload()"" onmouseover=""this.className='buttona';status='" & langRefreshMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/refresh.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langRefreshMenu & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=main.asp'"" onmouseover=""this.className='buttona';status='Home';"" onmouseout=""this.className='button';status='';""><img src=""images/document_home.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langDocumentsHome & """></span>"
    Response.Write "        <span class=""button"" onclick=""parent.fraTopic.focus();parent.fraTopic.print();"" onmouseover=""this.className='buttona';status='" & langPrintCurrDoc & "';"" onmouseout=""this.className='button';status='';""><img src=""images/print.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langPrintCurrDoc & """></span>"
    Response.Write "        &nbsp;"

    If HasPermission("CanEditDocuments") Then
     ' Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addfolder.asp'"" onmouseover=""this.className='buttona';status='" & langAddFolder & "';"" onmouseout=""this.className='button';status='';""><img src=""images/folder_closed.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddFolder & """></span>"
     ' Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addarticle.asp'"" onmouseover=""this.className='buttona';status='" & langAddDocument & "';"" onmouseout=""this.className='button';status='';""><img src=""images/document.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddDoc & """></span>"
    '  Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=addhelp.asp'"" onmouseover=""this.className='buttona';status='" & langAddHelp & "';"" onmouseout=""this.className='button';status='';""><img src=""images/helpdocument.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddHelp & """></span>"
       Response.Write "          <span class=""button"" onclick=""parent.fraToc.ToggleContextMenu();"" onmouseover=""this.className='buttona';status='" & langAddToggle & "';"" onmouseout=""this.className='button';status='';""><img src=""images/menu.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddToggle & """></span>"
      Response.Write "          &nbsp;"
      'JS 3/28/02 Response.Write "          <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=admin.asp'"" onmouseover=""this.className='buttona';status='" & langAdmin & "';"" onmouseout=""this.className='button';status='';""><img src=""images/admin.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAdmin & """></span>"
    End If
    Response.Write "      </td>"
  else
    Response.Write "      <td class=""submenu""><img src=""" & sRelDir & "images/spacer.gif"" height=""28"" width=""1"" border=""0""></td>"
  end if

  Response.Write "    </tr>"
  Response.Write "  </table>"
  Response.Write sTimeOutScript
End Sub


'----------------------------------------------------------------------------------------
' Sub DrawQuicklinks(searchCaption, dirLevel)
'----------------------------------------------------------------------------------------
Sub DrawQuicklinks(searchCaption, dirLevel)

  'Dim sDirLevel, i, sPath
  'sDirLevel = Left("../../../../../../", dirLevel*3)
  'sPath = Request.ServerVariables("SCRIPT_NAME")
%>
	<!-- Removed Quick links section as per Peters request 10/13/2005
  <div style="padding-bottom:8px;"><b><%=langQuicklinks%></b></div>
  <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/home_small.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>events"><%=langTabHome%></a></div>
  <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/document_home.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>docs"><%=langDocuments%></a></div>
  <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/newgroup.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>dirs"><%=langTabCommittees%></a></div>
 <!-- <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/newdisc.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>discussions"><%=langDiscussions%></a></div>
 <!-- <div class="quicklink">&nbsp;&nbsp;<img src="<%=sDirLevel%>images/newfav.gif" align="absmiddle">&nbsp;<a href="<%=sDirLevel%>favorites"><%=langFavorites%></a></div> //

  <% If searchCaption <> "" Then %>
   <form name="frmQlSearch" action="search.asp" method="get">
      <input type="hidden" name="p" value="1">
      <div style="padding-bottom:3px;"><%=searchCaption%>:</div>
      <input type="text" style="background-color:#eeeeee; border:1px solid #000000; width:144px;" name="s"><br>
      <div align="right"><a href="javascript:document.all.frmQlSearch.submit();"><img src="<%=sDirLevel%>images/go.gif" border="0"><%=langGo%></a>&nbsp;&nbsp;</div>
    </form>
  <% End If %>-->
<%
End Sub


'----------------------------------------------------------------------------------------
' Function AsciiToHtml( AsciiString )
'----------------------------------------------------------------------------------------
Function AsciiToHtml( AsciiString )
	Dim sTmp
	sTmp = AsciiString

	sTmp = Replace(sTmp, """", "&quot;")
	'sTmp = Replace(sTmp, "<", "&lt;")
	'sTmp = Replace(sTmp, ">", "&gt;")
	sTmp = Replace(sTmp, vbCrLf, "<br>")

	AsciiToHtml = sTmp
End Function


'----------------------------------------------------------------------------------------
' Function SQLText( DbString )
'----------------------------------------------------------------------------------------
Function SQLText( DbString )
	If DbString & "" = "" Then
		SQLText = NULL
	Else
		SQLText = Replace(DbString, "'", "''")
	End If
End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION DRAWADMINUSERS(sUSERID)
'--------------------------------------------------------------------------------------------------
Function DrawAdminUsers(sUserID)
	Dim sSql, oUsers, selected

	sSql = "SELECT userID, email, FirstName, LastName "
	sSql = sSql & " FROM Users "
	sSql = sSql & " WHERE orgId = " & session("orgId")
	sSql = sSql & " AND (IsRootAdmin is null or IsRootAdmin = 0) "
	sSql = sSql & " ORDER BY LastName, firstname "

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSql, Application("DSN"), 1, 3

	do while not oUsers.EOF
		if sUserID = oUsers("userID") then selected = " selected=""selected"" " else selected = ""
		response.write vbcrlf & vbtab & "<option value=""" & oUsers("userID") &"," & oUsers("FirstName") & " " & oUsers("LastName") & """" & selected & ">" & oUsers("FirstName") & " " & oUsers("LastName") & "</option>"
		oUsers.MoveNext
	Loop

	oUsers.close
	Set oUsers = Nothing 

End Function

'--------------------------------------------------------------------------------------------------
Function DrawAdminUsersNew(sUserID,isEmailRequired)
	Dim sSql, oUsers, selected

	sSql = "SELECT userID, FirstName, LastName "
	sSql = sSql & " FROM Users "
	sSql = sSql & " WHERE orgid = " & session("orgId")
	sSql = sSql & " AND (IsRootAdmin IS NULL OR IsRootAdmin = 0) "
    sSql = sSql & " AND isdeleted = 0 "

	if isEmailRequired = "Y" then
		sSql = sSql & " AND email IS NOT NULL "
		sSql = sSql & " AND email <> '' "
	end if

	 sSql = sSql & " ORDER BY LastName, firstname "

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSql, Application("DSN"), 1, 3

	while NOT oUsers.eof
   	if sUserID = oUsers("userID") then
       selected = " selected=""selected"""
    else
       selected = ""
    end if

  		response.write "  <option value=""" & oUsers("userID") & """" & selected & ">" & oUsers("FirstName") & " " & oUsers("LastName") & "</option>" & vbcrlf
  		oUsers.movenext
	wend

	oUsers.close
	Set oUsers = Nothing 

End Function

'------------------------------------------------------------------------------
Function DrawAdminUsersAssignedHideDeleted(sUserID)
	Dim sSql, oUsers, selected

 sSelectedUser                = 0
 sAssigned_notActive_userName = ""

	sSQL = "SELECT userID, "
 sSQL = sSQL & " email, "
 sSQL = sSQL & " FirstName, "
 sSQL = sSQL & " LastName, "
 sSQL = sSQL & " isdeleted "
	sSQL = sSQL & " FROM Users "
	sSQL = sSQL & " WHERE orgId = " & session("orgId")
	sSQL = sSQL & " AND (IsRootAdmin is null or IsRootAdmin = 0) "
 sSQL = sSQL & " AND isdeleted = 0 "
	sSQL = sSQL & " ORDER BY isdeleted, LastName, firstname "

	Set oUsers = Server.CreateObject("ADODB.Recordset")
	oUsers.Open sSQL, Application("DSN"), 1, 3

	do while not oUsers.EOF
    sSelectedValue = ""

    if sUserID <> "" then
     		if sUserID = oUsers("userID") then
          sSelectedValue = " selected=""selected"""
          sSelectedUser  = sSelectedUser + 1
       end if
    end if

  		response.write "<option value=""" & oUsers("userID") &"," & oUsers("FirstName") & " " & oUsers("LastName") & """" & sSelectedValue & ">" & oUsers("FirstName") & " " & oUsers("LastName") & "</option>" & vbcrlf

		oUsers.MoveNext
	loop

 if sSelectedUser = 0 then
    if sUserID <> "" then
       sAssigned_notActive_userName = GetAdminName(sUserID)
       sAssigned_notActive_userName = sAssigned_notActive_userName & " - Inactive"
    else
       sUserID = "0"
       sAssigned_notActive_userName = "No User Assignment Exists"
    end if

    response.write "<option value=""" & sUserID & "," & sAssigned_notActive_userName & """ style=""color:#ff0000;"" selected=""selected"">[" & sAssigned_notActive_userName & "]</option>" & vbcrlf
 end if

	oUsers.close
	Set oUsers = Nothing 

End Function

'--------------------------------------------------------------------------------------------------
function DrawDepartments(sDeptID,isEmailRequired)
 Dim sSql, oDepts, lcl_selected

 if sDeptID <> "" then
    lcl_groupIsInactive = checkGroupIsInactive(sDeptID)
 else
    lcl_groupIsInactive = False
 end if

	sSql = "SELECT groupid, orgid, groupname, groupdescription, 1 as queryorder "
 sSql = sSql & " FROM groups g "
 sSql = sSql & " WHERE g.orgid = " & session("orgid")
 sSql = sSql & " AND (SELECT count(ug.userid) "
 sSql = sSql &      " FROM usersgroups ug, users u "
 sSql = sSql &      " WHERE u.userid = ug.userid "
 sSql = sSql &      " AND ug.groupid = g.groupid "
 sSql = sSql &      " AND u.orgid = " & session("orgid")
 sSql = sSql &      " AND (isrootadmin IS NULL OR isrootadmin = 0) "

 if isEmailRequired = "Y" then
    sSql = sSql &   " AND u.email IS NOT NULL "
    sSql = sSql &   " AND u.email <> '' "
 end if

 sSql = sSql & ") > 0 "

 if lcl_groupIsInactive AND sDeptID <> "" then
    sSql = sSql & " UNION ALL "
    sSql = sSql & " SELECT g2.groupid, g2.orgid, g2.groupname + ' [inactive]' as groupname, g2.groupdescription, 2 as queryorder "
    sSql = sSql & " FROM groups g2 "
    sSql = sSql & " WHERE g2.orgid = " & session("orgid")
    sSql = sSql & " AND g2.groupid = " & sDeptID
    sSql = sSql & " AND g2.isInactive = 1 "
 end if

 sSql = sSql & " ORDER BY 5,3"

	set oDepts = Server.CreateObject("ADODB.Recordset")
	oDepts.Open sSql, Application("DSN"), 1, 3

 lcl_isInList = 0

	while not oDepts.eof
    if sDeptID <> "" then
       if CLng(sDeptID) = oDepts("groupid") then
          lcl_selected = " selected=""selected"""
          lcl_isInList = lcl_isInList + 1
       else
          lcl_selected = ""
          lcl_isInList = lcl_isInList
       end if
    else
       lcl_selected = ""
    end if

    response.write "  <option value=""" & oDepts("groupid") & """" & lcl_selected & ">" & oDepts("groupname") & "</option>" & vbcrlf

  		oDepts.movenext
	wend

	oDepts.close
	set oDepts = nothing

'1. Now check to see if the current DeptID is NOT in the list.
'2. If it is NOT then we have to still show the option so that the value is not lost when saving the request.
'3. The department not showing in the list may be because the admin viewing the request may not have access 
'  to the department that the request is assigned to.
'4. If this is the case then the request should NOT lose the department value because the user viewing it 
'   doesn't have access to the department assigned to the request.
 if lcl_isInList = 0 AND sDeptID <> "" then
    sSql = "SELECT groupid, groupname "
    sSql = sSql & " FROM groups "
    sSql = sSql & " WHERE groupid = " & sDeptID

    set oMissingDept = Server.CreateObject("ADODB.Recordset")
   	oMissingDept.Open sSql, Application("DSN"), 1, 3

    if not oMissingDept.eof then
       lcl_groupname = oMissingDept("groupname")
    else
       lcl_groupname = "[Department has been deleted]"
    end if

    oMissingDept.close
    set oMissingDept = nothing

    response.write "  <option value=""" & sDeptID & """ selected=""selected"">" & lcl_groupname & "</option>" & vbcrlf
 end if


end function

'------------------------------------------------------------------------------
function checkGroupIsInactive(p_groupid)
  lcl_return = False

  sSql = "SELECT isInactive FROM groups WHERE groupid = " & p_groupid

 	set oActive = Server.CreateObject("ADODB.Recordset")
	 oActive.Open sSql, Application("DSN"), 1, 3

  if not oActive.eof then
     lcl_return = oActive("isInactive")
  end if

  oActive.close
  set oActive = nothing

  checkGroupIsInactive = lcl_return

end function


'----------------------------------------------------------------------------------------
' FUNCTION SETORGANIZATIONPARAMETERS()
'----------------------------------------------------------------------------------------
Function SetOrganizationParameters()
	' SET DEFAULT RETURN VALUE
	iReturnValue = 1

	' BUILD CURRENT URL
	If request.servervariables("HTTPS") = "on" Then
		sProtocol = "https://"
	Else
		sProtocol = "http://"
	End If
	sSERVER = request.servervariables("SERVER_NAME")
	sCurrent = sProtocol & sSERVER & "/" & GetVirtualDirectyName()
	sCurrent = replace(sCurrent,"https:","http:")


	' LOOKUP CURRENT URL IN DATABASE
	sSql = "SELECT OrgID, OrgName, OrgPublicWebsiteURL, OrgEgovWebsiteURL, OrgTopGraphicLeftURL, OrgTopGraphicRightURL,  "
	sSql = sSql & "OrgWelcomeMessage, OrgActionLineDescription, OrgPaymentDescription, OrgHeaderSize, OrgTagline, "
	sSql = sSql & "OrgPaymentGateway, OrgActionOn, OrgPaymentOn, OrgDocumentOn, OrgCalendarOn, orgVirtualSiteName, "
	sSql = sSql & "orgRegistration, orgQueryTool, orgFaqOn, OrgFormLetterOn, OrgInternalEntry, gmtoffset, separate_index_catalog "
	sSql = sSql & "FROM Organizations O INNER JOIN TimeZones T ON O.OrgTimeZoneID = T.TimeZoneID WHERE O.OrgEgovWebsiteURL = '" & sCurrent & "'"

	Set oOrgInfo = Server.CreateObject("ADODB.Recordset")
	oOrgInfo.Open sSql, Application("DSN"), 0, 1
	
	If Not oOrgInfo.EOF Then
		iOrgID = oOrgInfo("OrgID")
		sOrgName = oOrgInfo("OrgName")
		Session("sOrgName") = sOrgName
		sHomeWebsiteURL = oOrgInfo("OrgPublicWebsiteURL")
		sEgovWebsiteURL = oOrgInfo("OrgEgovWebsiteURL")
		sTopGraphicLeftURL = oOrgInfo("OrgTopGraphicLeftURL")
		sTopGraphicRighURL = oOrgInfo("OrgTopGraphicRightURL")
		sWelcomeMessage = oOrgInfo("OrgWelcomeMessage")
		sActionDescription = oOrgInfo("OrgActionLineDescription")
		sPaymentDescription = oOrgInfo("OrgPaymentDescription")
		iHeaderSize = oOrgInfo("OrgHeaderSize")
		sTagline = oOrgInfo("OrgTagline")
		iPaymentGatewayID = oOrgInfo("OrgPaymentGateway")
		blnOrgAction = oOrgInfo("OrgActionOn")
		blnOrgPayment = oOrgInfo("OrgPaymentOn")
		blnOrgDocument = oOrgInfo("OrgDocumentOn")
		blnOrgCalendar = oOrgInfo("OrgCalendarOn")
		sorgVirtualSiteName = oOrgInfo("orgVirtualSiteName")
		blnOrgRegistration = oOrgInfo("orgRegistration")
		blnQuerytool = oOrgInfo("orgQueryTool")
		blnFaq = oOrgInfo("orgFaqOn")
		sOrgFormLetterOn = oOrgInfo("OrgFormLetterOn")
		sOrgInternalEntry = oOrgInfo("OrgInternalEntry")
		iTimeOffset =  oOrgInfo("gmtoffset")
		blnSeparateIndex = oOrgInfo("separate_index_catalog")
	Else
		iOrgID = 0
	End If
	Set oOrgInfo = Nothing 

	If Not IsNull(iOrgID) Then 
		iReturnValue = iOrgID
	Else
		iReturnValue = 0
	End If

	' RETURN VALUE
	SetOrganizationParameters = iReturnValue
	
End Function


'----------------------------------------------------------------------------------------
' GETVIRTUALDIRECTYNAME()
'----------------------------------------------------------------------------------------
Function GetVirtualDirectyName()

	sReturnValue = ""
	
	strURL = Request.ServerVariables("SCRIPT_NAME")
	strURL = Split(strURL, "/", -1, 0) 
	sReturnValue = "/" & strURL(1) 

	GetVirtualDirectyName = Replace(sReturnValue, "/", "")

End Function


'------------------------------------------------------------------------------
' Function DBsafe( strDB )
'  Make buffer Database 'safe'
'  Useful in building SQL Strings
'    strSQL="SELECT *....WHERE Value='" & DBSafe(strValue) & "';"
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' Function DBsafeWithHTML( strDB )
'------------------------------------------------------------------------------
Function DBsafeWithHTML( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
	End If 

	DBsafeWithHTML = sNewString
End Function


'----------------------------------------------------------------------------------------
' Function HTMLQuotes( strHTML )
'  Make buffer Quoted 'safe' 
'  Useful in building lines like:
'     "<INPUT TYPE=Hidden VALUE='" & HTMLQuotes(strValue) & "'>"
'----------------------------------------------------------------------------------------
Function HTMLQuotes( ByVal strHTML )
	HTMLQuotes = Replace( strHTML, Chr(38), "&#38;" )		' Ampersand
	HTMLQuotes = Replace( HTMLQuotes, Chr(34), "&#34;" )	' Double Quotes
	HTMLQuotes = Replace( HTMLQuotes, Chr(39), "&#39;" )	' Single Quotes
End Function

'------------------------------------------------------------------------------
Function Track_DBsafe( ByVal strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then Track_DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	Track_DBsafe = sNewString
End Function




'------------------------------------------------------------------------------
' FUNCTION  UserIsRootAdmin_old( iUserID )
'------------------------------------------------------------------------------
Function UserIsRootAdmin_old( ByVal iUserID )
	Dim sSql, oRoot, blnReturnValue

	' SET DEFAULT
	blnReturnValue = False

	' LOOKUP passed FEATURE FOR the current ORGANIZATION 
	sSql = "SELECT count(userid) as root_count FROM users "
	sSql = sSql & " WHERE isrootadmin = 1 and orgid = " & Session("OrgID") & " and userid = " & iUserID

	Set oRoot = Server.CreateObject("ADODB.Recordset")
	oRoot.Open  sSql, Application("DSN"), 3, 1
	
	If CLng(oRoot("root_count")) > 0 Then
		' the ORGANIZATION HAS the FEATURE
		blnReturnValue = True
	End If
	
	oRoot.close 
	Set oRoot = Nothing

	' set the RETURN  value
	UserIsRootAdmin = blnReturnValue
End Function 


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


'------------------------------------------------------------------------------
' Sub ShowHeader_old( sLevel )  - Do not use this, use ShowHeader instead
'------------------------------------------------------------------------------
Sub ShowHeader_old( sLevel )
	response.write "<table class=menu width=100% cellpadding=0 cellspacing=0 >"
	response.write "<tr><td width=360><img style=""padding:0px;"" src=""" & sLevel & "menu/logo2.jpg""></td><td align=RIGHT ><font style=""font-family: arial,tahoma; font-size: 11px; color:#FFFFFF;font-weight:bold;PADDING-RIGHT:15PX;"">ADMINISTRATION CONSOLE</font></td></tr>"
	response.write "</table>"
	response.write "<table style=""Height:30px;"" bgcolor=""#93bee1"" border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"">"
    response.write "<tr><td width=400>&nbsp;</td><td align=right><b>You are logged in as " & Session("FullName") & ".</b></td></tr>"
	response.write "<tr bgcolor=""#ffffff""><td height=""1"" colspan=""2""></td></tr>"
	response.write "<tr bgcolor=""#666666""><td height=""1"" colspan=""2""></td></tr>"
	response.write "</table>"
End Sub 


'------------------------------------------------------------------------------
' Sub ShowHeader( sLevel )
'------------------------------------------------------------------------------
Sub ShowHeader( ByVal sLevel )
	Dim oHeaderOrg

	Set oHeaderOrg = New classOrganization

	If session("orgid") = 0 Or session("orgid") = "" Then 
		oHeaderOrg.SetOrgId iOrgID
	End If 

	'Set up the org logo
	lcl_orgLogo = "<img src=""" & sLevel & oHeaderOrg.GetOrgDisplayName( "admin logo" ) & """ align=""left"" />"

	response.write "<style>" & vbcrlf
	response.write "  .supportHeaderTitle { font-family:arial; color:#ffff00; font-weight:bold; }" & vbcrlf
	response.write "  .supportHeaderLabel { font-family:arial; color:#ffff00; }" & vbcrlf
	response.write "</style>" & vbcrlf

	response.write "<table id=""pageheader"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbcrlf
	response.write "  <tr>" & vbcrlf
	'response.write "      <td width=""360""><img src=""" & sLevel & oHeaderOrg.GetOrgDisplayName( "admin logo" ) & """ /></td>" & vbcrlf

	'Determine if the org AND user have the permission to view the E-Gov Support Contact Info.
	If oHeaderOrg.orghasfeature("display_egovsupport_contactinfo") And userhaspermission(session("userid"), "display_egovsupport_contactinfo") Then 
		response.write "      <td>" & vbcrlf
		response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""color:#ffffff;"">" & vbcrlf
		response.write "            <tr valign=""top"">" & vbcrlf
		response.write "                <td rowspan=""2"">" & lcl_orgLogo & "</td>" & vbcrlf
		response.write "                <td align=""center"" colspan=""3"" class=""supportHeaderTitle hidetablet"">CONTACT E-GOV SUPPORT</td>" & vbcrlf
		response.write "            </tr>" & vbcrlf
		response.write "            <tr valign=""top"" class=""hidetablet"">" & vbcrlf
		response.write "                <td nowrap=""nowrap"" colspan=""2"">" & vbcrlf
		response.write "                    <span class=""supportHeaderLabel"">Phone: </span>513.591.7361<br />" & vbcrlf
		response.write "                    <span class=""supportHeaderLabel"">Email: </span><a href=""mailto:egovsupport@eclink.com?subject=Contact E-Gov Support (Admin Email)"" style=""color:#ffffff; text-decoration:underline;"">egovsupport@eclink.com</a>" & vbcrlf
		response.write "                </td>" & vbcrlf
		response.write "                <td nowrap=""nowrap"">" & vbcrlf
		response.write "                    <span class=""supportHeaderLabel"">Web Site: </span><a href=""http://www.egovlink.com/egovsupport/action.asp"" target=""_blank"" style=""color:#ffffff; text-decoration:underline;"">www.egovlink.com/egovsupport</a>" & vbcrlf
		response.write "                </td>" & vbcrlf
		response.write "            </tr>" & vbcrlf
		response.write "          </table>" & vbcrlf
		response.write "      </td>" & vbcrlf
	Else 
		response.write "      <td>" & lcl_orgLogo & "</td>" & vbcrlf
	End If 

	response.write "      <td align=""right"" id=""title"">ADMINISTRATION CONSOLE</td>" & vbcrlf
	response.write "  </tr>" & vbcrlf
	response.write "</table>" & vbcrlf
	response.write "<table id=""logheader"" border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbcrlf
	response.write "  <tr>" & vbcrlf
	response.write "      <td class=""namerow"" width=""400"">&nbsp;</td>" & vbcrlf

	If session("FullName") <> "" Then 
		response.write "      <td class=""namerow hidetablet"" align=""right"" ><strong>You are logged in as " & Session("FullName") & ".</strong></td>" & vbcrlf
	Else 
		response.write "      <td class=""namerow"" align=""right"">&nbsp;</td>" & vbcrlf
	End If 

	response.write "  </tr>" & vbcrlf
	response.write "</table>" & vbcrlf

	Set oHeaderOrg = Nothing 
End Sub 


'------------------------------------------------------------------------------
' Sub ShowDocumentsHeader( sLevel )
'------------------------------------------------------------------------------
Sub ShowDocumentsHeader( sLevel )
	response.write vbcrlf & "<table id=""pageheader"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
	response.write vbcrlf & "<tr><td width=""360""><img src=""" & sLevel & "menu/logo2.jpg""></td><td align=""right"" id=""title"" >ADMINISTRATION CONSOLE</td></tr>"
	response.write vbcrlf & "</table>"
	response.write vbcrlf & "<table id=""logheader"" border=""0"" cellpadding=""2"" cellspacing=""0"">"
    response.write vbcrlf & "<tr><td class=""namerow"" width=""175"">&nbsp;</td>"
	response.write vbcrlf & "<td class=""namerow"" valign=""top"">"
	response.write "      <table cellpadding=""0"" cellspacing=""0"" border=""0""><tr>"
	Response.Write "      <td class=""documentsubmenu""><img src=""images/spacer.gif"" width=""8"" height=""1""></td>"
	Response.Write "      <td class=""documentsubmenu"" width=""100%"" style=""padding-top:2px;"">&nbsp;"
'	Response.Write "        <span class=""button"" onclick=""toggleMenu();"" onmouseover=""this.className='buttona';status='" & langCollapseMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/collapse.gif"" id=""img_toggle"" width=""18"" height=""18"" border=""0"" alt=""" & langCollapseMenu & """></span>"
'	Response.Write "        <span class=""button"" onclick=""parent.fraToc.document.location.reload()"" onmouseover=""this.className='buttona';status='" & langRefreshMenu & "';"" onmouseout=""this.className='button';status='';""><img src=""images/refresh.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langRefreshMenu & """></span>"
'	Response.Write "        <span class=""button"" onclick=""parent.fraTopic.location.href='../load.asp?file=main.asp'"" onmouseover=""this.className='buttona';status='Home';"" onmouseout=""this.className='button';status='';""><img src=""images/document_home.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langDocumentsHome & """></span>"
'	Response.Write "        <span class=""button"" onclick=""parent.fraTopic.focus();parent.fraTopic.print();"" onmouseover=""this.className='buttona';status='" & langPrintCurrDoc & "';"" onmouseout=""this.className='button';status='';""><img src=""images/print.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langPrintCurrDoc & """></span>"
'	Response.Write "        <span class=""button"" onclick=""parent.fraToc.ToggleContextMenu();"" onmouseover=""this.className='buttona';status='" & langAddToggle & "';"" onmouseout=""this.className='button';status='';""><img src=""images/menu.gif"" width=""18"" height=""18"" border=""0"" alt=""" & langAddToggle & """></span>"
	Response.Write "      </td>"
	response.write "      </tr></table>"
	response.write vbcrlf & "</td>"
	response.write vbcrlf & "<td class=""namerow"" align=""right""><strong>You are logged in as " & Session("FullName") & ".</strong></td></tr>"
	response.write vbcrlf & "</table>"

End Sub 


'------------------------------------------------------------------------------
' Function UserHasPermission( iUserId, sFeature )
'------------------------------------------------------------------------------
Function UserHasPermission( iUserId, sFeature )
	' This will check that the org has the feature, handle rootadmin features, and checks user permissions - Call This
	Dim sSql, oPermission, bIsRootAdmin

	bIsRootAdmin = UserIsRootAdmin( iUserId )

	UserHasPermission = False 
	sSql = "SELECT F.haspermissions, F.rootadminrequired FROM egov_organization_features F, egov_organizations_to_features FO "
	sSql = sSql & " WHERE F.feature = '" & sFeature & "' AND FO.featureid = F.featureid AND FO.orgid = " & Session("orgid")

	Set oPermission = Server.CreateObject("ADODB.Recordset")
	oPermission.Open sSql, Application("DSN"), 0, 1

	If NOT oPermission.EOF Then
		If oPermission("rootadminrequired") Then 
			If bIsRootAdmin Then 
				' root admins get all these features by default
				UserHasPermission = True 
			Else 
				' check if permission can be assigned
				If oPermission("haspermissions") Then 
					' look up to see if they were assigned this feature
					UserHasPermission = UserHasFeature( iUserId, sFeature )
				End If 
			End If 
		Else 
			If oPermission("haspermissions") Then 
				' Needs permissions for user, so look them up
				UserHasPermission = UserHasFeature( iUserId, sFeature )
			Else
				' Permissions not required for this feature
				UserHasPermission = True 
			End If 
		End If 
	End If
	oPermission.close
	Set oPermission = Nothing 
	
End Function 


'------------------------------------------------------------------------------
' Function UserHasFeature( iUserId, sFeature )
'------------------------------------------------------------------------------
Function UserHasFeature( iUserId, sFeature )
	' Do not call this directly, call UserHasPermission instead
	Dim sSql, oPermission

	sSql = "SELECT COUNT(UF.Featureid) AS hits FROM egov_organization_features F, egov_users_to_features UF "
	sSql = sSql & " WHERE F.feature = '" & sFeature & "' AND UF.featureid = F.featureid AND UF.userid = " & iUserId

	Set oPermission = Server.CreateObject("ADODB.Recordset")
	oPermission.Open sSql, Application("DSN"), 0, 1

	oPermission.MoveFirst 
	If CLng(oPermission("hits")) > CLng(0) Then 
		UserHasFeature = True 
	Else
		UserHasFeature = False 
	End If 

	oPermission.Close
	Set oPermission = Nothing 
	
End Function 


'------------------------------------------------------------------------------
' integer GetUserPermissionLevel( iUserId, sFeature )
'------------------------------------------------------------------------------
Function GetUserPermissionLevel( ByVal iUserId, ByVal sFeature )
	' If you need permision level, check that they have permission first, then get the permission level
	Dim sSql, oPermissionLevel

	GetUserPermissionLevel = 0

	sSql = "SELECT ISNULL(permissionlevelid,0) AS permissionlevelid FROM egov_organization_features F, egov_users_to_features UF "
	sSql = sSql & " WHERE F.feature = '" & sFeature & "' AND UF.featureid = F.featureid AND UF.userid = " & iUserId

	Set oPermissionLevel = Server.CreateObject("ADODB.Recordset")
	oPermissionLevel.Open sSql, Application("DSN"), 0, 1

	If Not oPermissionLevel.EOF Then 
		oPermissionLevel.MoveFirst 
		GetUserPermissionLevel = oPermissionLevel("permissionlevelid")
	End If 

	oPermissionLevel.close
	Set oPermissionLevel = Nothing 

End Function 


'------------------------------------------------------------------------------
' integer GetPermissionLevelName( iPermissionLevel )
'------------------------------------------------------------------------------
Function GetPermissionLevelName( ByVal iPermissionLevelId )
	Dim sSql, oPermissionLevel

	GetPermissionLevelName = ""

	sSql = "Select permissionlevel From egov_feature_permission_levels "
	sSql = sSql & " Where permissionlevelid = " & iPermissionLevelId

	Set oPermissionLevel = Server.CreateObject("ADODB.Recordset")
	oPermissionLevel.Open sSql, Application("DSN"), 0, 1

	If Not oPermissionLevel.EOF Then 
		oPermissionLevel.MoveFirst 
		GetPermissionLevelName = oPermissionLevel("permissionlevel")
	End If 

	oPermissionLevel.close
	Set oPermissionLevel = Nothing 

End Function 


'------------------------------------------------------------------------------
' boolean isUserInDept( iUserId, iDeptId )
'------------------------------------------------------------------------------
Function isUserInDept( ByVal iUserId, ByVal iDeptId )
	Dim sSql, oDept

	isUserInDept = False 

	sSql = "Select count(groupid) as hits From usersgroups "
	sSql = sSql & " Where userid = " & iUserId & " and groupid = " & iDeptId

	Set oDept = Server.CreateObject("ADODB.Recordset")
	oDept.Open sSql, Application("DSN"), 3, 1

	If Not oDept.EOF Then 
		oDept.MoveFirst 
		If CLng(oDept("hits")) > CLng(0) Then 
			isUserInDept = True 
		End If 
	End If 

	oDept.close
	Set oDept = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetInternalDefaultContact( iOrgId )
'------------------------------------------------------------------------------
Function GetInternalDefaultContact( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(internal_default_contact,'') AS internal_default_contact "
	sSql = sSql & "FROM organizations WHERE orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetInternalDefaultContact = oRs("internal_default_contact")
	Else 
		GetInternalDefaultContact = "" 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' string GetInternalDefaultEmail( iOrgId )
'------------------------------------------------------------------------------
Function GetInternalDefaultEmail( ByVal iOrgId )
	Dim sSql, oDefaultEmail

	'GetInternalDefaultEmail = "webmaster@eclink.com" 
	GetInternalDefaultEmail = "noreply@eclink.com"

	' This will yield one row per org with the default email for the internal contact
	sSql = "Select isnull(internal_default_email,'') as internal_default_email from organizations where orgid = " & iOrgId 

	Set oDefaultEmail = Server.CreateObject("ADODB.Recordset")
	oDefaultEmail.Open sSql, Application("DSN"), 0, 1

	If Not oDefaultEmail.EOF Then 
		GetInternalDefaultEmail = oDefaultEmail("internal_default_email")
	End If 

	oDefaultEmail.close
	Set oDefaultEmail = Nothing 

End Function 


'------------------------------------------------------------------------------
' boolean OrgHasNeighborhoods( iOrgId )
'------------------------------------------------------------------------------
Function OrgHasNeighborhoods( ByVal iOrgId )
	Dim sSql, oNeighborhood

	sSql = "SELECT count(neighborhoodid) as hits FROM egov_neighborhoods where orgid = " & iorgid 
	
	Set oNeighborhood = Server.CreateObject("ADODB.Recordset")
	oNeighborhood.Open sSql, Application("DSN"), 0, 1

	If clng(oNeighborhood("hits")) > 0 Then
		OrgHasNeighborhoods = True 
	Else
		OrgHasNeighborhoods = False 
	End if
	
	oNeighborhood.close
	Set oNeighborhood = Nothing 

End Function 


'------------------------------------------------------------------------------
' void DisplayNeighborhoods( iorgid, iNeighborhoodId )
'------------------------------------------------------------------------------
Sub DisplayNeighborhoods( ByVal iorgid, ByVal iNeighborhoodId )
	Dim sSql, oNeighborhood 

	sSql = "SELECT neighborhoodid, neighborhood FROM egov_neighborhoods where orgid = " & iorgid & " order by neighborhood"

	Set oNeighborhood = Server.CreateObject("ADODB.Recordset")
	oNeighborhood.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<select name=""egov_users_neighborhoodid"">"	
	response.write vbcrlf &  "<option value=""0"""
	If clng(iNeighborhoodId) = clng(0) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Not on List...</option>"
		
	Do While NOT oNeighborhood.EOF 
		response.write vbcrlf & "<option value=""" &  oNeighborhood("neighborhoodid") & """"
		If clng(iNeighborhoodId) = clng(oNeighborhood("neighborhoodid")) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oNeighborhood("neighborhood") & "</option>"
		oNeighborhood.MoveNext
	Loop

	response.write vbcrlf & "</select>"

	oNeighborhood.close
	Set oNeighborhood = Nothing 

End Sub  
 

'------------------------------------------------------------------------------
' Function GetDefaultRelationShipId( iOrgid )
'------------------------------------------------------------------------------
Function GetDefaultRelationShipId( iOrgid )
	Dim sSql, oRelationShip

	sSql = "SELECT relationshipid FROM egov_familymember_relationships where orgid = " & iorgid & " and isdefault = 1"
	
	Set oRelationShip = Server.CreateObject("ADODB.Recordset")
	oRelationShip.Open sSql, Application("DSN"), 3, 1

	If Not oRelationShip.EOF Then
		GetDefaultRelationShipId = oRelationShip("relationshipid") 
	Else
		GetDefaultRelationShipId = 0 
	End if
	
	oRelationShip.close
	Set oRelationShip = nothing
End Function 


'------------------------------------------------------------------------------
' Sub UpdateFamilyId( iUserId, iFamilyId, iRelationshipId, iNeighborhoodid, bResidencyVerified )
'------------------------------------------------------------------------------
Sub UpdateFamilyId( iUserId, iFamilyId, iRelationshipId, iNeighborhoodid, bResidencyVerified, sEmailnotavailable )
	Dim sSql, oCmd

	sSql = "Update egov_users set familyid = " & iFamilyId
	
	If iRelationshipId <> "" Then 
		sSql = sSql & ", relationshipid = " & iRelationshipId
	Else
		sSql = sSql & ", relationshipid = NULL"
	End If 

	If iNeighborhoodid <> "" Or iNeighborhoodid <> "0" Then 
		sSql = sSql & ", neighborhoodid = " & iNeighborhoodid
	Else
		sSql = sSql & ", neighborhoodid = NULL"
	End If 

	If bResidencyVerified Then
		sSql = sSql & ", residencyverified = 1"
	Else
		sSql = sSql & ", residencyverified = 0"
	End If 

	sSql = sSql & ", emailnotavailable = " & sEmailnotavailable

	sSql = sSql & " where userid = " & iUserId 

'	response.write sSql
'	response.End 

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing
End Sub 


'------------------------------------------------------------------------------
' integer iUserId = GetFamilyId( iUserId )
'------------------------------------------------------------------------------
Function GetFamilyId( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT familyid FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFamilyId = oRs("familyid")
	Else
		GetFamilyId = iUserId
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' string GetAdminName( iUserId )
'------------------------------------------------------------------------------
Function GetAdminName( ByVal iUserId )
	Dim sSql, oRs

	If iUserID <> "" Then
	   	sSql = "SELECT firstname + ' ' + lastname AS username FROM users WHERE userid = " & iUserId 

	   	Set oRs = Server.CreateObject("ADODB.Recordset")
	   	oRs.Open sSql, Application("DSN"), 0, 1

	   	If Not oRs.EOF Then
			GetAdminName = Trim(oRs("username"))
		Else
			GetAdminName = ""
		End If 

	   	oRs.Close
	   	Set oRs = Nothing 
	Else
	 	GetAdminName = ""
	End If

End Function 


'------------------------------------------------------------------------------
' void GetAdminFirstAndLastName oRs("adminuserid"), sFirstName, sLastName 
'------------------------------------------------------------------------------
Sub GetAdminFirstAndLastName( ByVal iUserId, ByRef sFirstName, ByRef sLastName )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname FROM users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		sFirstName = Trim(oRs("firstname"))
		sLastName = Trim(oRs("lastname"))
	Else
		sFirstName = ""
		sLastName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' string GetAdminPhone( iUserId ) 
'------------------------------------------------------------------------------
Function GetAdminPhone( ByVal iUserId ) 
	Dim sSql, oRs

	sSql = "SELECT businessnumber FROM users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetAdminPhone = FormatPhoneNumber(oRs("businessnumber"))
	Else
		GetAdminPhone = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' Function GetAdminLocation( iAdminLocationId )
'------------------------------------------------------------------------------
Function GetAdminLocation( ByVal iAdminLocationId )
	Dim sSql, oRs

	sSql = "SELECT name FROM egov_class_location WHERE locationid = " & iAdminLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetAdminLocation = oRs("name")
	Else
		GetAdminLocation = ""
	End If 
	
	oRs.Close 
	Set oRs = Nothing
End Function 


'------------------------------------------------------------------------------
' Function GetCitizenName( iUserId )
'------------------------------------------------------------------------------
Function GetCitizenName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT userfname + ' ' + userlname AS username FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetCitizenName = oRs("username")
	Else
		GetCitizenName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' Function GetFamilyOwnerName( iFamilyId )
'------------------------------------------------------------------------------
Function GetFamilyOwnerName( ByVal iFamilyId )
	Dim sSql, oRs

	sSql = "SELECT userfname + ' ' + userlname AS familyname FROM egov_users WHERE useremail IS NOT NULL AND familyid = " & iFamilyId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFamilyOwnerName = oRs("familyname")
	Else
		GetFamilyOwnerName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' Function GetFamilyEmail( iUserId )
'------------------------------------------------------------------------------
Function GetFamilyEmail( ByVal iUserId )
	Dim sSql, oRs, iFamilyId

	iFamilyId = GetFamilyId( iUserId )

	sSql = "SELECT useremail FROM egov_users WHERE useremail IS NOT NULL AND headofhousehold = 1 AND familyid = " & iFamilyId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFamilyEmail = oRs("useremail")
	Else
		GetFamilyEmail = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' Function UserHasFamilyMembers( iUserID, iFamilyId ) 
'------------------------------------------------------------------------------
Function UserHasFamilyMembers( ByVal iUserID, ByVal iFamilyId ) 
	Dim sSql, oRs

	sSql = "SELECT COUNT(userid) AS hits FROM egov_users WHERE familyid = " & iFamilyId & " AND userid != " & iUserID

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If CLng(oRs("hits")) > CLng(0) Then
		UserHasFamilyMembers = True 
	Else
		UserHasFamilyMembers = False 
	End if
	
	oRs.Close
	Set oRs = Nothing 

End Function


'------------------------------------------------------------------------------
' Function GetCurrentURL()
'------------------------------------------------------------------------------
Function GetCurrentURL()
	Dim prot, https, domainname, filename, querystring, sURL

	prot = "http" 
	https = lcase(request.ServerVariables("HTTPS")) 
	If https <> "off" Then
		prot = "https" 
	End If 
	domainname = Request.ServerVariables("SERVER_NAME") 
	filename = Request.ServerVariables("SCRIPT_NAME") 
	querystring = Request.ServerVariables("QUERY_STRING") 
	sURL =  prot & "://" & domainname & filename 
	If querystring <> "" Then 
		sURL = sURL & "?" & querystring 
	End If 

	GetCurrentURL = sURL
End Function 


'------------------------------------------------------------------------------
' Function GetCitizenAge( dBirthDate )
'------------------------------------------------------------------------------
Function GetCitizenAge( ByVal dBirthDate )
	Dim iMonths, iAge

	If Not IsNull(dBirthDate) And IsDate(dBirthDate) Then 
		iMonths = DateDiff("m", dBirthDate, Now())
		If iMonths = 0 Then 
			iMonths = 1 
		End If 
		iAge = FormatNumber(iMonths / 12, 1)
	Else
		iAge = 21
	End If 
	
	GetCitizenAge = iAge
End Function 


'------------------------------------------------------------------------------
' Function CitizenHasRecreationActivities( iUserId )
'------------------------------------------------------------------------------
Function CitizenHasRecreationActivities( ByVal iUserId )
	Dim sSql, oRs

	CitizenHasRecreationActivities = False 

	sSql = "SELECT COUNT(classlistid) AS hits FROM egov_familymembers F, egov_class_list L "
	sSql = sSql & " WHERE F.familymemberid = L.familymemberid AND F.userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If clng(oRs("hits")) > 0 Then
		CitizenHasRecreationActivities = True 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string sFormatString = FormatForJavaScript( sString )
'------------------------------------------------------------------------------
Function FormatForJavaScript( ByVal sString )
	Dim sNewString 

	sNewString = Replace( sString, "'","\'" )

	FormatForJavaScript = sNewString
End Function 


'------------------------------------------------------------------------------
' double dAccountBalance = GetCitizenCurrentBalance( iUserId )
'------------------------------------------------------------------------------
Function GetCitizenCurrentBalance( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountbalance,0.00) AS accountbalance FROM egov_users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetCitizenCurrentBalance = CDbl(oRs("accountbalance"))
	Else
		GetCitizenCurrentBalance = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' Function Proper( ByVal sString )
'------------------------------------------------------------------------------
Function Proper( ByVal sString )
	Dim sNewString

	sNewString = Trim(sString)
	If Len(sNewString) > 0 Then 
		sNewString = UCase(Left(sNewString,1)) & Mid(sNewString,2)
	End If 

	Proper = sNewString
End Function 


'------------------------------------------------------------------------------
' Function GetUserPageSize( iUserId )
'------------------------------------------------------------------------------
Function GetUserPageSize( iUserId )
	Dim sSql, oPage
	
	sSql = "Select isnull(pagesize,20) as pagesize from users where userid = " & iUserId

	Set oPage = Server.CreateObject("ADODB.Recordset")
	oPage.Open  sSql, Application("DSN"), 3, 1

	If Not oPage.EOF Then 
		GetUserPageSize = clng(oPage("pagesize"))
	Else
		GetUserPageSize = 20 ' This is the default
	End If 

	oPage.close
	Set oPage = Nothing 

End Function 

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
' string GetOrgDisplay( iorgid, sDisplay )
'------------------------------------------------------------------------------
Function GetOrgDisplay( ByVal iOrgId, ByVal sDisplay )
	Dim sSql, oDisplay

	' SET DEFAULT
	GetOrgDisplay = ""

	' LOOKUP passed Display FOR the passed Organization 
	sSql = "SELECT ISNULL(OD.displaydescription, D.displaydescription) AS displaydescription "
	sSql = sSql & " FROM egov_organizations_to_displays OD, egov_organization_displays D "
	sSql = sSql & " WHERE OD.displayid = D.displayid "
	sSql = sSql & " AND orgid = " & iOrgId
	sSql = sSql & " AND D.display = '" & sDisplay & "' "

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open sSql, Application("DSN"), 3, 1

	If Not oDisplay.EOF Then
		' the ORGANIZATION HAS the Display
		GetOrgDisplay = oDisplay("displaydescription")
	End If

	oDisplay.Close 
	Set oDisplay = Nothing

End Function


'------------------------------------------------------------------------------
' string GetDisplayName( iDisplayId )
'------------------------------------------------------------------------------
Function GetDisplayName( ByVal iDisplayId )
	Dim sSql, oDisplay

	' SET DEFAULT
	GetDisplayName = ""

	' LOOKUP passed Display FOR the passed Organization 
	sSql = "SELECT displayname FROM egov_organization_displays WHERE displayid = " & iDisplayId 

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSql, Application("DSN"), 3, 1
	
	If Not oDisplay.EOF Then
		' the ORGANIZATION HAS the Display
		GetDisplayName = oDisplay("displayname")
	End If
	
	oDisplay.close 
	Set oDisplay = Nothing

End Function

'------------------------------------------------------------------------------
' string GetOrgDisplayWithId( iorgid, iDisplayId, bUsesDisplayName )
'------------------------------------------------------------------------------
Function GetOrgDisplayWithId( ByVal iOrgId, ByVal iDisplayId, ByVal bUsesDisplayName )
	Dim sSql, oDisplay, sField

	' SET DEFAULT
	GetOrgDisplayWithId = ""
	If bUsesDisplayName Then
		sField = "displayname"
	Else
		sField = "displaydescription"
	End If 

	' LOOKUP passed Display FOR the passed Organization 
	sSql = "SELECT ISNULL(OD." & sField & ", D." & sField & ") AS displayfield "
	sSql = sSql & " FROM egov_organizations_to_displays OD, egov_organization_displays D "
	sSql = sSql & " WHERE OD.displayid = D.displayid AND orgid = " & iOrgId & " AND D.displayid = " & iDisplayId

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSql, Application("DSN"), 3, 1
	
	If Not oDisplay.EOF Then
		' the ORGANIZATION HAS the Display
		GetOrgDisplayWithId = oDisplay("displayfield")
	End If
	
	oDisplay.close 
	Set oDisplay = Nothing

End Function


'------------------------------------------------------------------------------
' integer GetDisplayId( sDisplay )
'------------------------------------------------------------------------------
Function GetDisplayId( ByVal sDisplay )
	Dim sSql, oDisplay

	sSql = "SELECT displayid FROM egov_organization_displays WHERE display = '" & sDisplay & "' "

	Set oDisplay = Server.CreateObject("ADODB.Recordset")
	oDisplay.Open  sSql, Application("DSN"), 0, 1
	
	If Not oDisplay.EOF Then
		GetDisplayId = clng(oDisplay("displayid"))
	Else
		GetDisplayId = clng(0)
	End If
	
	oDisplay.Close 
	Set oDisplay = Nothing

End Function 


'----------------------------------------------------------------------------------------
' Function GetOrgName()
'----------------------------------------------------------------------------------------
Function GetOrgName( ByVal iorgid )
	Dim sSql, oRs

	sSql = "Select orgname FROM organizations WHERE orgid = " & iorgid
'		response.write sSql
'		response.end

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetOrgName = oRs("orgname")
	End If
		
	oRs.Close
	Set oRs = Nothing

End Function 


'----------------------------------------------------------------------------------------
' Function GetOrgValue()
'----------------------------------------------------------------------------------------
Function GetOrgValue( ByVal sOrgField )
	Dim sSql, oRs

	sSql = "SELECT " & sOrgField & " AS orgvalue FROM organizations WHERE orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetOrgValue = oRs("orgvalue")
	End If
		
	oRs.Close
	Set oRs = Nothing
End Function 


'------------------------------------------------------------------------------
' Function GetUserLocation( iLocationId )
'------------------------------------------------------------------------------
Function GetUserLocation( ByVal iLocationId )
	Dim sSql, oLocation
	
	sSql = "Select name from egov_class_location where locationid = " & iLocationId

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open  sSql, Application("DSN"), 3, 1

	If Not oLocation.EOF Then 
		GetUserLocation = oLocation("name")
	Else
		GetUserLocation = ""
	End If 

	oLocation.close
	Set oLocation = Nothing 

End Function 


'------------------------------------------------------------------------------
' Function GetUserLocationId( iUserID )
'------------------------------------------------------------------------------
Function GetUserLocationId( ByVal iUserID )
	Dim sSql, oLocation
	
	sSql = "Select isnull(locationid,0) as locationid from users where userid = " & iUserID

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open  sSql, Application("DSN"), 3, 1

	If Not oLocation.EOF Then 
		GetUserLocationId = clng(oLocation("locationid"))
	Else
		GetUserLocationId = 0
	End If 

	oLocation.close
	Set oLocation = Nothing 
End Function 


'------------------------------------------------------------------------------
' Sub ShowUserLocations( )
'------------------------------------------------------------------------------
Sub ShowUserLocations( )
	Dim sSql, oLocation
	
	sSql = "Select locationid, name from egov_class_location where orgid = " & session("orgid") & " order by name"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open  sSql, Application("DSN"), 3, 1

	If Not oLocation.EOF Then 
		response.write vbcrlf & "<select name=""locationid"" onChange=""SetLocation();"">"
		If clng(0) = clng(session("locationid")) Then ' They do not have one assigned
			response.write vbcrlf & "<option value=""0"" selected=""selected"" >Select a Location...</option>"
		End If 
		Do While Not oLocation.EOF 
			response.write vbcrlf & "<option value=""" & oLocation("locationid") & """"
			If clng(oLocation("locationid")) = clng(session("locationid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oLocation("name") & "</option>"
			oLocation.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oLocation.close
	Set oLocation = Nothing 

End Sub 


'------------------------------------------------------------------------------
' Function getDefaultOrgValue( sColumn )
'------------------------------------------------------------------------------
Function getDefaultOrgValue( ByVal sColumn )
	Dim sSql, oValue
	
	sSql = "Select isnull(" & sColumn & ",'') as defaultvalue from organizations where orgid = " & session("orgid")

	Set oValue = Server.CreateObject("ADODB.Recordset")
	oValue.Open  sSql, Application("DSN"), 3, 1

	If Not oValue.EOF Then 
		getDefaultOrgValue = oValue("defaultvalue")
	Else
		getDefaultOrgValue = ""
	End If 

	oValue.close
	Set oValue = Nothing 
End Function 


'------------------------------------------------------------------------------
' ShowFamilyAccounts iUserID 
'------------------------------------------------------------------------------
Sub ShowFamilyAccounts( ByVal iUserID )
	Dim sSql, oRs

	sSql = "SELECT userfname, userlname, userid, ISNULL(accountbalance,0.00) AS accountbalance "
	sSql = sSql & " FROM egov_users WHERE familyid = " & GetFamilyId( iUserId )
	sSql = sSql & " ORDER BY userlname, userfname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""accountid"">"
		'response.write vbcrlf & "<option value=""0"">Select Account...</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("userid") & """>" & oRs("userfname") & " " & oRs("userlname") & " (" & FormatNumber(oRs("accountbalance"),2) & ") " & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' Function GetItemTypeId( sType )
'------------------------------------------------------------------------------
Function GetItemTypeId( ByVal sType )
	Dim sSql, oEntry, sTypeId

	sSql = "Select itemtypeid from egov_item_types Where itemtype = '" & sType & "'"

	Set oEntry = Server.CreateObject("ADODB.Recordset")
	oEntry.Open sSql, Application("DSN"), 0, 1

	If Not oEntry.EOF Then 
		sTypeId = oEntry("itemtypeid") 
	Else 
		sTypeId = 0
	End If 

	oEntry.close
	Set oEntry = Nothing

	GetItemTypeId = sTypeId
End Function 


'------------------------------------------------------------------------------
' Function GetItemType( iTypeId )
'------------------------------------------------------------------------------
Function GetItemType( ByVal iTypeId )
	Dim sSql, oEntry, sType

	sSql = "Select itemtype from egov_item_types Where itemtypeid = " & iTypeId 

	Set oEntry = Server.CreateObject("ADODB.Recordset")
	oEntry.Open sSql, Application("DSN"), 0, 1

	If Not oEntry.EOF Then 
		sType = oEntry("itemtype") 
	Else 
		sType = ""
	End If 

	oEntry.close
	Set oEntry = Nothing

	GetItemType = sType
End Function 


'------------------------------------------------------------------------------
' Function GetmaxPaymentTypeId()
'------------------------------------------------------------------------------
Function GetmaxPaymentTypeId( ByVal iOrgId )
	Dim sSql, oMax

	sSql = "Select max(paymenttypeid) as maxid from egov_organizations_to_paymenttypes Where orgid = " & iOrgId

	Set oMax = Server.CreateObject("ADODB.Recordset")
	oMax.Open sSql, Application("DSN"), 0, 1

	GetmaxPaymentTypeId = clng(oMax("maxid"))

	oMax.close
	Set oMax = Nothing 
	
End Function 


'------------------------------------------------------------------------------
' InsertPaymentInformation iPaymentId, iLedgerId, iPaymentTypeId, sAmount, sStatus, sCheckNo, iAccountId
'------------------------------------------------------------------------------
Sub InsertPaymentInformation( ByVal iPaymentId, ByVal iLedgerId, ByVal iPaymentTypeId, ByVal sAmount, ByVal sStatus, ByVal sCheckNo, ByVal iAccountId )
	Dim oCmd, sSql 

	sSql = "INSERT INTO egov_verisign_payment_information (paymentid, ledgerid, paymenttypeid, amount, "
	sSql = sSql & "paymentstatus, checkno, citizenuserid) Values (" & iPaymentid & ", " & iLedgerId & ", " 
	sSql = sSql & iPaymentTypeId & ", " & sAmount & ", '" & sStatus & "', " & sCheckNo & ", " & iAccountId & " )"
'	response.write sSql & "<br /><br />"

	RunSQLStatement sSql

'	Set oCmd = Server.CreateObject("ADODB.Command")
'	With oCmd
'		.ActiveConnection = Application("DSN")
'		.CommandText = sSql
'		.Execute
'	End With
'	Set oCmd = Nothing

End Sub 


'------------------------------------------------------------------------------
' Function HasCitizensAccounts( iPaymentTypeId )
'------------------------------------------------------------------------------
Function HasCitizensAccounts( ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT requirescitizenaccount FROM egov_paymenttypes WHERE paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If oRs("requirescitizenaccount") Then
		HasCitizensAccounts = True 
	Else
		HasCitizensAccounts = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' Function HasChecks( iPaymentTypeId )
'------------------------------------------------------------------------------
Function HasChecks( ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT requirescheckno FROM egov_paymenttypes WHERE paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If oRs("requirescheckno") Then
		HasChecks = True 
	Else
		HasChecks = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' AdjustCitizenAccountBalance iUserID, sEntryType, sAmount 
'------------------------------------------------------------------------------
Sub AdjustCitizenAccountBalance( ByVal iUserID, ByVal sEntryType, ByVal sAmount )
	Dim sNewBalance, cPriorBalance, sSql

	cPriorBalance = GetCitizenCurrentBalance( iUserId )

	If sEntryType = "credit" Then
		sNewBalance = CDbl(cPriorBalance) + CDbl(sAmount)
	Else  ' debit
		sNewBalance = CDbl(cPriorBalance) - CDbl(sAmount)
	End If 

	sSql = "UPDATE egov_users SET accountbalance = " & sNewBalance & " WHERE userid = " & iUserID

	RunSQLStatement sSql

End Sub 


'------------------------------------------------------------------------------
' integer GetPaymentAccountId( iOrgId, iPaymentTypeId )
'------------------------------------------------------------------------------
Function GetPaymentAccountId( ByVal iOrgId, ByVal iPaymentTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountid,0) AS accountid FROM egov_organizations_to_paymenttypes "
	sSql = sSql & "WHERE orgid = " & iOrgId & " AND paymenttypeid = " & iPaymentTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetPaymentAccountId = CLng(oRs("accountid"))
	Else
		GetPaymentAccountId = CLng(0) 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' ShowAccountPicks iAccountid, sIdField 
'------------------------------------------------------------------------------
Sub ShowAccountPicks( ByVal iAccountid, ByVal sIdField )
	Dim sSql, oRs

	sSql = "SELECT accountid, accountname FROM egov_accounts "
	sSql = sSql & " WHERE accountstatus = 'A' AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY accountname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""accountid" & sIdField & """>"

		If CLng(iAccountid) = CLng(0) Then 
			response.write vbcrlf & "<option value=""0"" selected=""selected"" >Select an Account</option>"
		end if

		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("accountid") & """ "  

			If CLng(iAccountid) = CLng(oRs("accountid")) Then 
				response.write " selected=""selected"" "
			End If 

			response.write ">" & oRs("accountname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write "</select>" & vbcrlf
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' string sRefundName = GetRefundName()
'------------------------------------------------------------------------------
Function GetRefundName( )
	Dim sSql, oRs

	sSql = "SELECT T.paymenttypename FROM egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE T.isrefundmethod = 1 AND T.paymenttypeid = O.paymenttypeid AND O.orgid = " & Session("OrgID") 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRefundName = oRs("paymenttypename")
	Else
		GetRefundName = "Refund Voucher"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
'  integer iPaymentTypeId = GetRefundPaymentTypeId()
'------------------------------------------------------------------------------
Function GetRefundPaymentTypeId( )
	Dim sSql, oRs

	sSql = "SELECT T.paymenttypeid FROM egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & " WHERE T.isrefundmethod = 1 AND T.paymenttypeid = O.paymenttypeid AND O.orgid = " & Session("OrgID") 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetRefundPaymentTypeId = oRs("paymenttypeid")
	Else
		GetRefundPaymentTypeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
'  Function OrgHasRefundDebitAcct( )
'------------------------------------------------------------------------------
Function OrgHasRefundDebitAcct( )
	Dim sSql, oRefund

	sSql = "Select count(O.paymenttypeid) as hits From egov_organizations_to_paymenttypes O, egov_paymenttypes P "
	sSql = sSql & " Where P.isrefunddebit = 1 and P.paymenttypeid = O.paymenttypeid and O.orgid = " & Session("OrgID") 

	Set oRefund = Server.CreateObject("ADODB.Recordset")
	oRefund.Open sSql, Application("DSN"), 0, 1

	If clng(oRefund("hits")) > clng(0) Then
		OrgHasRefundDebitAcct = True 
	Else
		OrgHasRefundDebitAcct =  False 
	End If 

	oRefund.Close
	Set oRefund = Nothing 
End Function 


'------------------------------------------------------------------------------
'  Function GetRefundDebitAccountId( )
'------------------------------------------------------------------------------
Function GetRefundDebitAccountId( )
	Dim sSql, oRefund

	sSql = "select isnull(O.accountid,0) as accountid from egov_paymenttypes T, egov_organizations_to_paymenttypes O "
	sSql = sSql & " where T.isrefunddebit = 1 and T.paymenttypeid = O.paymenttypeid and O.orgid = " & Session("OrgID") 

	Set oRefund = Server.CreateObject("ADODB.Recordset")
	oRefund.Open sSql, Application("DSN"), 0, 1

	If Not oRefund.EOF Then
		GetRefundDebitAccountId = oRefund("accountid")
	Else
		GetRefundDebitAccountId = 0
	End If 

	oRefund.Close
	Set oRefund = Nothing 
End Function 


'------------------------------------------------------------------------------
' Function MakeRefundLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iIsCCRefund )
'------------------------------------------------------------------------------
Function MakeRefundLedgerEntry( iOrgID, iAccountId, iJournalId, cAmount, iItemTypeId, sEntryType, sPlusMinus, iItemId, iIsPaymentAccount, iPaymentTypeId, cPriorBalance, iIsCCRefund )
	Dim sSql, oInsert, iLedgerId

	iLedgerId = 0

	sSql = "Insert Into egov_accounts_ledger ( paymentid, orgid, entrytype, accountid, amount, itemtypeid, plusminus, "
	sSql = sSql & "itemid, ispaymentaccount, paymenttypeid, priorbalance, isccrefund ) Values ( "
	sSql = sSql & iJournalId & ", " & iOrgID & ", '" & sEntryType & "', " & iAccountId & ", " & cAmount & ", " & iItemTypeId & ", '" & sPlusMinus & "', " 
	sSql = sSql & iItemId & ", " & iIsPaymentAccount & ", " & iPaymentTypeId & ", " & cPriorBalance & ", " & iIsCCRefund & " )"
	sSql = "SET NOCOUNT ON;" & sSql & ";SELECT @@IDENTITY AS ROWID;"
'	response.write sSql & "<br /><br />"
'	response.End 
	session("sSql") = sSql

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3
	session("sSql") = ""

	iLedgerId = oInsert("ROWID")

	oInsert.close
	Set oInsert = Nothing

	MakeRefundLedgerEntry = iLedgerId

End Function 


'------------------------------------------------------------------------------
' Function FormatPhoneNumber( Number )
'------------------------------------------------------------------------------
Function FormatPhoneNumber( ByVal Number )
	If Len(Number) = 10 Then
		FormatPhoneNumber = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhoneNumber = Number
	End If
End Function


'------------------------------------------------------------------------------
' Sub BreakOutAddress( ByVal sAddress, ByRef sStreetNumber, ByRef sStreetName )
'------------------------------------------------------------------------------
Sub BreakOutAddress( ByVal sAddress, ByRef sStreetNumber, ByRef sStreetName )
  	'Break out the number from the name, should be at first space from left
   	Dim iPos
   	iPos = InStr(sAddress, " ")
   	If Not IsNull(iPos) And iPos > 0 Then
     		sStreetNumber = Trim(Left(sAddress, (iPos - 1)))
     		If IsNumeric(sStreetNumber) Then 
       			sStreetName = Trim(Mid(sAddress,(iPos + 1)))
     		Else
      			'The first field is not a number, so this is just a street name or something else
       			sStreetNumber = ""
       			sStreetName   = sAddress
     		End If 
   	Else
    		'no space so maybe this is a street name or something that will not work
     		sStreetNumber = ""
     		sStreetName   = sAddress
   	End If 
End Sub 


'------------------------------------------------------------------------------
' Function IsValidAddress( sStreetNumber, sStreetName )
'------------------------------------------------------------------------------
Function IsValidAddress( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oAddress

'	Old way to validate
'	sSql = "Select count(residentaddressid) as hits From egov_residentaddresses "
'	sSql = sSql & " Where residentstreetnumber = '" & dbsafe(sStreetNumber) & "' and residentstreetname = '" & dbsafe(sStreetName) & "' and orgid = " & Session("OrgID") 

'	New way to validate as of 4/3/2008
	sSql = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses "
	sSql = sSql & " WHERE residentstreetnumber = '" & dbsafe(sStreetNumber) & "' "
	sSql = sSql & " AND (ltrim(rtrim(residentstreetname)) = '" & dbsafe(sStreetName) & "' "
	sSql = sSql & " OR ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) = '" & dbsafe(sStreetName) & "' "
	sSql = sSql & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) = '" & dbsafe(sStreetName) & "' "
	sSql = sSql & " OR ltrim(rtrim(residentstreetprefix)) + ' ' + ltrim(rtrim(residentstreetname)) + ' ' + ltrim(rtrim(streetsuffix)) + ' ' + ltrim(rtrim(streetdirection)) = '" & dbsafe(sStreetName) & "' )"
	sSql = sSql & " AND orgid = " & Session("OrgID")  

	Set oAddress = Server.CreateObject("ADODB.Recordset")
	oAddress.Open sSql, Application("DSN"), 0, 1

	If clng(oAddress("hits")) > clng(0) Then
		IsValidAddress = True 
	Else
		IsValidAddress =  False 
	End If 

	oAddress.Close
	Set oAddress = Nothing 

End Function

'--------------------------------------------------------------------------------------------------
function IsValidAddress_byStreetName(p_orgid, sStreetName)
  dim sSQL, oAddress

  lcl_return = False

 'New way to validate as of 4/3/2008
 	sSQL = "SELECT COUNT(residentaddressid) AS hits FROM egov_residentaddresses "
 	sSQL = sSQL & " WHERE (residentstreetname = '" & track_dbsafe(sStreetName) & "' "
 	sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = '" & track_dbsafe(sStreetName) & "' "
 	sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = '" & track_dbsafe(sStreetName) & "' "
 	sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = '" & track_dbsafe(sStreetName) & "' )"
 	sSQL = sSQL & " AND orgid = " & p_orgid

 	set oAddress = Server.CreateObject("ADODB.Recordset")
 	oAddress.Open sSQL, Application("DSN"), 0, 1

 	if CLng(oAddress("hits")) > CLng(0) then
   		lcl_return = True
 	end if

 	oAddress.close
	 set oAddress = nothing 

  IsValidAddress_byStreetName = lcl_return

end function

'------------------------------------------------------------------------------
' Function GetMembershipDesc( iPoolPassId )
'------------------------------------------------------------------------------
Function GetMembershipDesc( iPoolPassId )
	Dim sSql, oRs

	sSql = "SELECT membershipdesc FROM egov_poolpasspurchases P, egov_poolpassrates R, egov_memberships M "
	sSql = sSql & " WHERE P.rateid = R.rateid AND R.membershipid = M.membershipid and P.poolpassid = " & iPoolPassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetMembershipDesc = oRs("membershipdesc")
	Else
		GetMembershipDesc = "Pool"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' Function isValidEmail( sEmail )
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
' Function GetFixtureFeeMethod( iOrgid )
'------------------------------------------------------------------------------
Function GetFixtureFeeMethod( ByVal iOrgid )
	Dim sSql, oRs

	sSql = "SELECT permitfeemethodid FROM egov_permitfeemethods WHERE isfixture = 1 and orgid = " & iOrgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetFixtureFeeMethod = CLng(oRs("permitfeemethodid"))
	Else
		GetFixtureFeeMethod = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
' Function GetValuationFeeMethod( iOrgid )
'------------------------------------------------------------------------------
Function GetValuationFeeMethod( ByVal iOrgid )
	Dim sSql, oRs

	sSql = "SELECT permitfeemethodid FROM egov_permitfeemethods WHERE isvaluation = 1 and orgid = " & iOrgid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetValuationFeeMethod = CLng(oRs("permitfeemethodid"))
	Else
		GetValuationFeeMethod = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'------------------------------------------------------------------------------
function changeBGColor(p_bgcolor,p_first_color,p_second_color)
 'Set up the row colors
  if p_first_color <> "" then
     lcl_first_color = p_first_color
  else
     lcl_first_color = "#eeeeee"
  end if

  if p_second_color <> "" then
     lcl_second_color = p_second_color
  else
     lcl_second_color = "#ffffff"
  end if

 'Determine which row color to display
  if p_bgcolor <> "" then
     lcl_bgcolor = p_bgcolor

     if UCASE(lcl_bgcolor) = UCASE(lcl_first_color) then
        lcl_bgcolor = lcl_second_color
     else
        lcl_bgcolor = lcl_first_color
     end if
  else
     lcl_bgcolor = lcl_first_color
  end if

  changeBGColor = lcl_bgcolor

end function

'------------------------------------------------------------------------------
function isFeatureOffline(p_feature)
 'If the user is ROOT ADMIN then bypass the check for any features that may be offline
  if UserIsRootAdmin(session("userid")) then
     lcl_feature_offline = "N"
  else
     sSql = "SELECT distinct f.feature_offline "
     sSql = sSql & " FROM egov_organization_features f "
     sSql = sSql & " WHERE f.feature_offline = 'Y' "
     sSql = sSql & " AND UPPER(f.feature) IN ('" & UCASE(REPLACE(p_feature,",","','")) & "') "

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSql, Application("DSN"),3,1

     if not rs.eof then
        if rs("feature_offline") <> "Y" then
           lcl_feature_offline = "N"
        else
           lcl_feature_offline = "Y"
        end if
     else
        lcl_feature_offline = "N"
     end If
     rs.close
	 Set rs = Nothing 
  end if

  isFeatureOffline = lcl_feature_offline

end function


'------------------------------------------------------------------------------
' Function ParentFeatureIsOffline( p_feature )
'	Similar to isFeatureOffline(p_feature) but you can pass the child feature of the page
'------------------------------------------------------------------------------
Function ParentFeatureIsOffline( p_feature )
	Dim sSql, oRs 

	'If the user is ROOT ADMIN then bypass the check for any features that may be offline
	If UserIsRootAdmin( session("userid") ) Then 
		ParentFeatureIsOffline = False 
	Else 
		sSql = "SELECT UPPER(B.feature_offline) AS feature_offline FROM egov_organization_features A, egov_organization_features B "
		sSql = sSql & " WHERE A.parentfeatureid = B.featureid AND A.feature = '" & p_feature & "'"

		set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.EOF Then
			If oRs("feature_offline") = "Y" Then 
				ParentFeatureIsOffline = True 
			Else
				ParentFeatureIsOffline = False 
			End If 
		Else 
			ParentFeatureIsOffline = False 
		End  If 
		oRs.Close
		Set oRs = Nothing 
	End If 

End Function 



'------------------------------------------------------------------------------
' Sub PageDisplayCheck( sFeature, sLevel )
'------------------------------------------------------------------------------
Sub PageDisplayCheck( ByVal sFeature, ByVal sLevel )
	' This sub handles all page permission and feature offline functions from one point. 

	If Not UserHasPermission( Session("UserId"), sFeature ) and  Not UserHasPermission( Session("UserId"), replace(sFeature, "permit","permitv2" )) and  Not UserHasPermission( Session("UserId"), sFeature & "v2" ) Then
		response.redirect sLevel & "permissiondenied.asp"
	End If

	If ParentFeatureIsOffline( sFeature ) Then 
		response.redirect sLevel & "admin/outage_feature_offline.asp"
	End If 

End Sub 


'------------------------------------------------------------------------------
function buildStreetAddress(sStreetNumber, sPrefix, sStreetName, sSuffix, sDirection)
  lcl_street_name = ""

  if trim(sStreetNumber) <> "" then
     lcl_street_name = trim(sStreetNumber)
  end if

  if trim(sPrefix) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sPrefix)
     else
        lcl_street_name = trim(sPrefix)
     end if
  end if

  if trim(sStreetName) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sStreetName)
     else
        lcl_street_name = trim(sStreetName)
     end if
  end if

  if trim(sSuffix) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sSuffix)
     else
        lcl_street_name = trim(sSuffix)
     end if
  end if

  if trim(sDirection) <> "" then
     if lcl_street_name <> "" then
        lcl_street_name = lcl_street_name & " " & trim(sDirection)
     else
        lcl_street_name = trim(sDirection)
     end if
  end if

  buildStreetAddress = trim(lcl_street_name)

end function

'------------------------------------------------------------------------------
sub GetAddressInfoNew(ByVal iOrgHasFeature_LargeAddressList, ByVal p_orgid, ByVal sStreetNumber, ByVal sStreetName, _
                      ByRef sNumber, ByRef sPrefix, ByRef sAddress, ByRef sSuffix, ByRef sDirection, _
                      ByRef sLatitude, ByRef sLongitude, ByRef sCity, ByRef sState, ByRef sZip, ByRef sCounty, _
                      ByRef sParcelID, ByRef sListedOwner, ByRef sLegalDescription, ByRef sResidentType, _
                      ByRef sRegisteredUserID, ByRef sValidStreet)

 'Set return variables
  sValidStreet      = "N"
  lcl_streetnumber  = dbsafe(sStreetNumber)
  lcl_streetname    = dbsafe(sStreetName)
		sNumber           = ""
  sPrefix           = ""
  sAddress          = ""
  sSuffix           = ""
  sDirection        = ""
	 sLatitude         = ""
  sLongitude        = ""
  sCity             = ""
  sState            = ""
  sZip              = ""
  sCounty           = ""
  sParcelID         = ""
  sListedOwner      = ""
  sLegalDescription = ""
  sResidentType     = ""
  sRegisteredUserID = ""

	 sSQL = "SELECT residentstreetnumber, residentstreetprefix, residentstreetname, streetsuffix, streetdirection, "
  sSQL = sSQL & " isnull(latitude,0.00) as latitude, isnull(longitude,0.00) as longitude, residentcity, residentstate, residentzip, "
  sSQL = sSQL & " county, parcelidnumber, listedowner, legaldescription, residenttype, registereduserid "
  sSQL = sSQL & " FROM egov_residentaddresses "

  if iOrgHasFeature_LargeAddressList then
    'Format the streetname
     lcl_streetname = "'" & lcl_streetname & "'"

   	 sSQL = sSQL & " WHERE orgid = " & p_orgid
     sSQL = sSQL & " AND excludefromactionline = 0 "
     sSQL = sSQL & " AND UPPER(residentstreetnumber) = UPPER('" & lcl_streetnumber & "') "
     sSQL = sSQL & " AND (residentstreetname = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetprefix + ' ' + residentstreetname = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetname + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " OR residentstreetname + ' ' + streetsuffix + ' ' + streetdirection = " & lcl_streetname
     sSQL = sSQL & " )"
  else
    'Format the streetname
     if lcl_streetname = "" then
        lcl_streetname = 0
     end if

    	sSQL = sSQL & " WHERE residentaddressid = " & lcl_streetname
  end if

	 set oAddress = Server.CreateObject("ADODB.Recordset")
	 oAddress.Open sSQL, Application("DSN"), 3, 1
	
	 if not oAddress.eof then
   		sNumber           = trim(oAddress("residentstreetnumber"))
     sPrefix           = oAddress("residentstreetprefix")
   		sAddress          = oAddress("residentstreetname")
     sSuffix           = oAddress("streetsuffix")
     sDirection        = oAddress("streetdirection")
	 	  sLatitude         = oAddress("latitude")
   		sLongitude        = oAddress("longitude")
     sCity             = oAddress("residentcity")
     sState            = oAddress("residentstate")
     sZip              = oAddress("residentzip")
     sCounty           = oAddress("county")
     sParcelID         = oAddress("parcelidnumber")
     sListedOwner      = oAddress("listedowner")
     sLegalDescription = oAddress("legaldescription")
     sResidentType     = oAddress("residenttype")
     sRegisteredUserID = oAddress("registereduserid")
     sValidStreet      = "Y"
	 end if

	 oAddress.close
	 set oAddress = nothing

end sub

'------------------------------------------------------------------------------
function getPrinter_CardLayout(p_orgid)

  if p_orgid <> "" then

     sSql = "SELECT p.layoutid "
     sSql = sSql & " FROM organizations o, egov_membershipcard_printers p "
     sSql = sSql & " WHERE orgid = " & p_orgid
     sSql = sSql & " AND o.membershipcard_printer = p.printerid "

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSql, Application("DSN"),3,1

     if not rs.eof then
        getPrinter_CardLayout = rs("layoutid")
     else
        getPrinter_CardLayout = 0
     end if
  else
     getPrinter_CardLayout = 0
  end if

end function

'------------------------------------------------------------------------------
function formatCardDisplayValue(p_value)
  if UCASE(p_value) = "[NO VALUE]" then
     lcl_return = REPLACE(UCASE(p_value),"[NO VALUE]","")
  else
     lcl_return = p_value
  end if

  formatCardDisplayValue = lcl_return

end function

'--------------------------------------------------------------------------------------------------
Function GetTimeOffset( ByVal iOrgID )
	Dim sSql, oTime

	sSql = "SELECT T.gmtoffset FROM Organizations O, TimeZones T WHERE O.OrgTimeZoneID = T.TimeZoneID AND O.orgid = " & iOrgID

	Set oTime = Server.CreateObject("ADODB.Recordset")
	oTime.Open sSql, Application("DSN"), 0, 1

	If Not oTime.EOF Then 
  		GetTimeOffset =  clng(oTime("gmtoffset"))
	Else
  		GetTimeOffset = clng(0)
	End If 

	oTime.Close
	Set oTime = Nothing 

End Function

'------------------------------------------------------------------------------
function ConvertDateTimetoTimeZone()
  lcl_return     = ""
  datCurrentDate = Now()

 'Get the local date/time for the timezone
  sSql = "SELECT dbo.GetLocalDate(" & session("orgid") & ", '" & datCurrentDate & "') AS localDate "
  sSql = sSql & " FROM organizations "
  sSql = sSql & " WHERE orgid = " & session("orgid")

 	set oTimeOffset = Server.CreateObject("ADODB.Recordset")
	 oTimeOffset.Open sSql, Application("DSN"), 3, 1

  if not oTimeOffset.eof then
     lcl_localdate = oTimeOffset("localDate")
  else
     lcl_localdate = "NULL"
  end if

  oTimeOffset.close
  set oTimeOffset = nothing

  'datGMTDateTime = DateAdd("h",5,datCurrentDate)
  'iTimeOffset    = GetTimeOffset( Session("OrgID") )
  'lcl_return     = DateAdd("h",iTimeOffset,datGMTDateTime)

  lcl_return = lcl_localdate

  ConvertDateTimetoTimeZone = lcl_return

end function

'------------------------------------------------------------------------------
function dbready_string( ByVal p_value, ByVal p_length )
  lcl_return = ""
  if p_value <> "" AND p_length <> "" then
     lcl_return = trim(p_value)
     lcl_return = replace(lcl_return,"<","&lt;")
     lcl_return = replace(lcl_return,">","&gt;")

    'Verify the length
     if len(lcl_return) > p_length then
        lcl_return = mid(lcl_return,1,p_length)
     end if

     lcl_return = replace(lcl_return,"'","''")

  end if

  dbready_string = lcl_return

end function

'------------------------------------------------------------------------------
function dbready_date( ByVal p_value )
  lcl_return = False

  if p_value <> "" then
'     lcl_return = trim(p_value)

     if isDate(p_value) then
        lcl_return = True
     end if
  end if

  dbready_date = lcl_return

end function

'------------------------------------------------------------------------------
function dbready_number( ByVal p_value )
  lcl_return = False

  if p_value <> "" then
'     lcl_return = trim(p_value)

     if isNumeric(p_value) then
        lcl_return = True
     end if
  end if

  dbready_number = lcl_return

end function

'-------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSqli = "INSERT INTO my_table_dtb (notes) VALUES ('" & replace(p_value,"'","''") & "')"
  set rsi = Server.CreateObject("ADODB.Recordset")
 	rsi.Open sSqli, Application("DSN"), 3, 1
end Sub


'-------------------------------------------------------------------------------------------------
' Sub SetupMessagePopUp( sMessage )
'	To get this to work, put a call to this in the SCRIPT section of the page, put a call to
'	DisplayMessagePopUp below the footer and include scripts/modules.js in the page.
'-------------------------------------------------------------------------------------------------
Sub SetupMessagePopUp( sMessage )
	response.write vbcrlf & "<div id=""successmessage"">"
		response.write vbcrlf & "<div id=""successmessageheader"">"
			response.write vbcrlf & "<span class=""successmessagetext"">Success</span>"
		response.write vbcrlf & "</div>"
		response.write vbcrlf & "<div id=""displayedmessage"">"
			response.write vbcrlf & "<span class=""successmessagetext"">" & sMessage & "</span>"
		response.write vbcrlf & "</div>"
	response.write vbcrlf & "</div>"
End Sub 


'-------------------------------------------------------------------------------------------------
' Sub DisplayMessagePopUp()
'-------------------------------------------------------------------------------------------------
Sub DisplayMessagePopUp()
	response.write vbcrlf & "var msgw = (getWindowWidth() - 300)/2;"
	response.write vbcrlf & "var msgh = (getWindowHeight() - 150)/2;"
	response.write vbcrlf & "Event.observe(window, 'load', function() {  "
	response.write vbcrlf & "$(""successmessage"").style.left = msgw + 'px';"
	response.write vbcrlf & "$(""successmessage"").style.top = msgh + 'px';"
	response.write vbcrlf & "$(""successmessage"").style.visibility = 'visible';"
	response.write vbcrlf & "fadeout.delay(.50);        "
	response.write vbcrlf & "Element.hide.delay(1.5, ""successmessage"");    "
	response.write vbcrlf & "});    "
	response.write vbcrlf & "function fadeout()"
	response.write vbcrlf & "{        "
	response.write vbcrlf & "new Effect.Opacity(""successmessage"", {duration:1.0, from:1.0, to:0.0});    "
	response.write vbcrlf & "}   "
End Sub 


'-------------------------------------------------------------------------------------------------
' integer iIdentity = RunInsertStatement( sInsertStatement )
'-------------------------------------------------------------------------------------------------
Function RunInsertStatement( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush
	session("InsertSQL") = sInsertStatement

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.Open sSql, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.Close
	Set oInsert = Nothing

	RunInsertStatement = iReturnValue
	session("InsertSQL") = ""

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

'------------------------------------------------------------------------------
function checkSendEmail(iDoNotSendSelfEmail,iDoNotSendAllEmail,iUserCompareEmail)

  lcl_return = "Y"

 'Check to see if the user does not want to send the email to self/all assigned
  lcl_doNotSendSelfEmail = iDoNotSendSelfEmail
  lcl_doNotSendAllEmail  = iDoNotSendAllEmail

 'Check to see if the "do not send to all" is clicked
  if lcl_doNotSendAllEmail = "on" then
     lcl_return = "N"
  end if

 'Check to see if the "do not send to self" is clicked
  if lcl_doNotSendSelfEmail = "on" then
     lcl_sessionuser_email = getUserEmail(session("userid"))

     if UCASE(lcl_sessionuser_email) = UCASE(iUserCompareEmail) then
        lcl_return = "N"
     end if
  end if

  checkSendEmail = lcl_return

end function


'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
sub displayHelpIcon(iHelpDoc)
  lcl_helpdoc     = ""
  lcl_onclick     = ""
  lcl_onmouseover = ""
  lcl_onmouseout  = ""


  if iHelpDoc <> "" then
     sSql = "SELECT d.documenturl "
     sSql = sSql & " FROM egov_helpdocs h, documents d "
     sSql = sSql & " WHERE h.documentid = d.documentid "
     sSql = sSql & " AND UPPER(helpdoc_name) = '" & UCASE(iHelpDoc) & "'"

     set oHelpDoc = Server.CreateObject("ADODB.Recordset")
     oHelpDoc.Open sSql, Application("DSN"),3,1

     if not oHelpDoc.eof then
	lcl_helpdoc = oHelpDoc("documenturl")
        'lcl_helpdoc = replace(lcl_helpdoc,"/public_documents300","")
        lcl_helpdoc = replace(lcl_helpdoc,session("sitename"),"egovsupport")
     end if

     oHelpDoc.close
     set oHelpDoc = nothing

    'If documentation exists then display the Help Icon
     if lcl_helpdoc <> "" then

       'All help documentation is stored under E-Gov Support!
        'lcl_helpdoc_url = "http://www.egovlink.com/egovsupport/admin"
        lcl_helpdoc_url = lcl_helpdoc_url & lcl_helpdoc

        'lcl_onmouseover = " onmouseover=""tooltip.show('CLICK FOR HELP');"""
        'lcl_onmouseout  = " onmouseout=""tooltip.hide();"""
        lcl_onmouseover = ""
        lcl_onmouseout  = ""
        lcl_onclick     = " onclick=""location.href='" & lcl_helpdoc_url & "'"""

        response.write "<img src=""../images/help.jpg"" border=""0"" style=""cursor:pointer""" & lcl_onmouseover & lcl_onmouseout & lcl_onclick & " />" & vbcrlf
     end if
  end if

end sub

'------------------------------------------------------------------------------
sub sendToRSS(iFeedID, p_orgid, iRssRowID, iRssTitle, iRssDesc, iRssLink, iRssPubDate, iRssCreatedByID, iCreatedByUser)

  sSql = "INSERT INTO egov_rss (feedid, orgid, rowid, title, description, rsslink, publicationdate, createdbyid, createdbyname) VALUES ("
  sSql = sSql &       iFeedID                        & ", "
  sSql = sSql &       p_orgid                        & ", "
  sSql = sSql &       iRssRowID                      & ", "
  sSql = sSql & "'" & DBsafeWithHTML(iRssTitle)      & "', "
  sSql = sSql & "'" & DBsafeWithHTML(iRssDesc)       & "', "
  sSql = sSql & "'" & DBsafeWithHTML(iRssLink)       & "', "
  sSql = sSql & "'" & DBsafeWithHTML(iRssPubDate)    & "', "
  sSql = sSql &       iRssCreatedByID                & ", "
  sSql = sSql & "'" & DBsafeWithHTML(iCreatedByUser) & "');"

 	set oSendRSS = Server.CreateObject("ADODB.Recordset")
	 oSendRSS.Open sSql, Application("DSN"), 3, 1

  set oSendRSS = nothing

 'Now update the "lastbuilddate" for the feed.
  sSql = "UPDATE egov_rssfeeds SET lastbuilddate = '" & DBsafeWithHTML(iRssPubDate) & "' "
  sSql = sSql & " WHERE feedid = " & iFeedID

 	set oUpdateFeed = Server.CreateObject("ADODB.Recordset")
	 oUpdateFeed.Open sSql, Application("DSN"), 3, 1

  set oUpdateFeed = nothing

end sub

'------------------------------------------------------------------------------
function getFeedNameByOrgFeature(p_feature)
  lcl_return = ""

  if p_feature <> "" then
     sSql = "SELECT CL_portaltype "
     sSql = sSql & " FROM egov_organization_features "
     sSql = sSql & " WHERE UPPER(feature) = '" & UCASE(p_feature) & "' "

    	set oFeedName = Server.CreateObject("ADODB.Recordset")
   	 oFeedName.Open sSql, Application("DSN"), 3, 1

     if not oFeedName.eof then
        lcl_return = oFeedName("CL_portaltype")
     end if

     oFeedName.close
     set oFeedName = nothing
  end if

  getFeedNameByOrgFeature = lcl_return

end function

'------------------------------------------------------------------------------
function getFeedIDByFeedName(p_feedname)
  lcl_return = 0

  if p_feedname <> "" then
     sSql = "SELECT feedid "
     sSql = sSql & " FROM egov_rssfeeds "
     sSql = sSql & " WHERE UPPER(feedname) = '" & UCASE(p_feedname) & "' "

    	set oFeedID = Server.CreateObject("ADODB.Recordset")
   	 oFeedID.Open sSql, Application("DSN"), 3, 1

     if not oFeedID.eof then
        lcl_return = oFeedID("feedid")
     end if

     oFeedID.close
     set oFeedID = nothing
  end if

  getFeedIDByFeedName = lcl_return

end function

'------------------------------------------------------------------------------
function checkRSSLogExists(iOrgID, iRowID, iFeedName)
  lcl_return = False

  if iOrgID <> "" AND iRowID <> "" AND iFeedName <> "" then
     sSql = "SELECT count(r.rssid) as total_rss "
     sSql = sSql & " FROM egov_rss r, egov_rssfeeds f "
     sSql = sSql & " WHERE r.feedid = f.feedid "
     sSql = sSql & " AND UPPER(f.feedname) = '" & iFeedName & "' "
     sSql = sSql & " AND r.orgid = " & iOrgID
     sSql = sSql & " AND r.rowid = " & iRowID

     set oRSSCount = Server.CreateObject("ADODB.Recordset")
   	 oRSSCount.Open sSql, Application("DSN"), 3, 1

     if oRSSCount("total_rss") > 0 then
        lcl_return = True
     end if

     oRSSCount.close
     set oRSSCount = nothing
  end if

  checkRSSLogExists = lcl_return

end function


'------------------------------------------------------------------------------
' CreatePaymentsControlRow( sLogEntry, sFeature )
'------------------------------------------------------------------------------
Function CreatePaymentsControlRow( ByVal sLogEntry, ByVal sFeature )
	Dim sSql, iPaymentControlNumber

	sSql = "INSERT INTO paymentlog ( orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & session("orgid") & ", 'Admin', " & sFeature & ", '" & dbready_string(sLogEntry,500) & "' )"
	'response.write sSql & "<br /><br />"

	iPaymentControlNumber = RunInsertStatement( sSql )

	sSql = "UPDATE paymentlog SET paymentcontrolnumber = " & iPaymentControlNumber
	sSql = sSql & " WHERE paymentlogid = " & iPaymentControlNumber
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

	CreatePaymentsControlRow = iPaymentControlNumber

End Function 


'------------------------------------------------------------------------------
' AddToPaymentsLog( iPaymentControlNumber, sLogEntry, sFeature )
'------------------------------------------------------------------------------
Sub AddToPaymentsLog( ByVal iPaymentControlNumber, ByVal sLogEntry, ByVal sFeature  )
	Dim sSql

	sSql = "INSERT INTO paymentlog ( paymentcontrolnumber, orgid, applicationside, feature, logentry ) VALUES ( "
	sSql = sSql & iPaymentControlNumber & ", " & session("orgid") & ", 'Admin', " & sFeature & ", '" & dbready_string(sLogEntry,500) & "' )"
	'response.write sSql & "<br /><br />"
	RunSQLStatement( sSql )

End Sub


'------------------------------------------------------------------------------
sub displayActionLineForms(p_orgid, p_value, p_showblankline)

  if p_showblankline = "Y" then
     response.write "  <option value=""""></option>" & vbcrlf
  end if

 	sSql = "SELECT action_form_id, action_form_name "
  sSql = sSql & " FROM dbo.egov_form_list_200 "
  sSql = sSql & " WHERE form_category_id <> 6 "
  sSql = sSql & " AND action_form_internal <> 1 "
  sSql = sSql & " AND orgid=" & p_orgid
  sSql = sSql & " ORDER BY UPPER(action_form_name) "

  set oALForms = Server.CreateObject("ADODB.Recordset")
  oALForms.Open sSql, Application("DSN"), 3, 1

  if not oALForms.eof then
     do while not oALForms.eof

        if p_value = oALForms("action_form_id") then
           lcl_selected_formid = " selected=""selected"""
        else
           lcl_selected_formid = ""
        end if

        response.write "  <option value=""" & oALForms("action_form_id") & """" & lcl_selected_formid & ">" & oALForms("action_form_name") & "</option>" & vbcrlf

        oALForms.movenext
     loop
  end if

  oALForms.close
  set oALForms = nothing

end sub

'------------------------------------------------------------------------------
function getActionLineFormName(p_orgid, p_action_form_id)

  lcl_return = ""

  if p_action_form_id <> "" then
    	sSql = "SELECT action_form_name "
     sSql = sSql & " FROM dbo.egov_form_list_200 "
     sSql = sSql & " WHERE form_category_id <> 6 "
     sSql = sSql & " AND action_form_internal <> 1 "
     sSql = sSql & " AND orgid=" & p_orgid
     sSql = sSql & " AND action_form_id = " & p_action_form_id

     set oALFormName = Server.CreateObject("ADODB.Recordset")
     oALFormName.Open sSql, Application("DSN"), 3, 1

     if not oALFormName.eof then
        lcl_return = oALFormName("action_form_name")
     end if

     oALFormName.close
     set oALFormName = nothing
  end if

  getActionLineFormName = lcl_return

end function

'------------------------------------------------------------------------------
sub displayAddThisButton()

 '-----------------------------------------------------------------------------
 'Must have the following .js file in your code.
  '<script type=""text/javascript"" src=""https://s7.addthis.com/js/200/addthis_widget.js""></script>

 'Put this at the top of your <SCRIPT> tag.
  'var addthis_pub="cschappacher";
 '-----------------------------------------------------------------------------

 'AddThis Button
  response.write "<a href=""http://www.addthis.com/bookmark.php?v=20"" onmouseover=""return addthis_open(this, '', '[URL]', '[TITLE]')"" onmouseout=""addthis_close()"" onclick=""return addthis_sendto()"">" & vbcrlf
  response.write "<img src=""http://s7.addthis.com/static/btn/lg-addthis-en.gif"" width=""125"" height=""16"" alt=""Bookmark and Share"" style=""border:0"" />" & vbcrlf
  response.write "</a>" & vbcrlf

end sub

'------------------------------------------------------------------------------
sub checkForRSSFeed(p_orgid, p_featureid, p_feedname, p_egovwebsiteurl)

  if p_egovwebsiteurl <> "" then
     lcl_egovwebsiteurl = p_egovwebsiteurl & "/rssfeeds.asp"
  else
     lcl_egovwebsiteurl = "#"
  end if

  if p_featureid <> "" OR p_feedname <> "" then

    'Determine if this feed exists and is active.
     lcl_exists = "N"

     sSql = "SELECT distinct 'Y' as feedexists "
     sSql = sSql & " FROM egov_rssfeeds rf, egov_organizations_to_features FO, egov_organization_features f "
     'sSql = sSql & " WHERE UPPER(rss.feature) = UPPER(f.feature) "
     'sSql = sSql & " AND F.featureid = FO.featureid "
     sSql = sSql & " WHERE rf.feedfeatureid = f.featureid "
     sSql = sSql & " AND rf.feedfeatureid = fo.featureid "
     sSql = sSql & " AND FO.orgid = " & p_orgid

     if p_featureid <> "" then
        sSql = sSql & " AND upper(rf.feature) = (select upper(of2.feature) "
        sSql = sSql &                          " from egov_organization_features of2 "
        sSql = sSql &                          " where of2.featureid = " & p_featureid & ") "
     else
        sSql = sSql & " AND UPPER(rf.feedname) = '" & UCASE(p_feedname) & "'"
     end if

     sSql = sSql & " AND (select count(rss.rssid) "
     sSql = sSql &      " from egov_rss rss "
     sSql = sSql &      " where rss.orgid = " & p_orgid
     sSql = sSql &      " and rss.feedid = rf.feedid) > 0 "

     set oFeedExists = Server.CreateObject("ADODB.Recordset")
     oFeedExists.Open sSql, Application("DSN"), 3, 1

     if not oFeedExists.eof then
        lcl_exists = oFeedExists("feedexists")
     end if

     oFeedExists.close
     set oFeedExists = nothing

    'If the feed exists then display the image/link.
     if lcl_exists = "Y" then
        response.write "<a href="""  & lcl_egovwebsiteurl & """>" & vbcrlf
        response.write "<img src=""" & p_egovwebsiteurl & "images/socialsites/icon_rss.png"" border=""0"" alt=""Subscribe to this RSS Feed"" />" & vbcrlf
        response.write "</a>" & vbcrlf
     end if

  end if

end sub

'------------------------------------------------------------------------------
function getFeatureID( ByVal p_feature )
  lcl_return = 0

  if p_feature <> "" then
     sSql = "SELECT featureid "
     sSql = sSql & " FROM egov_organization_features "
     sSql = sSql & " WHERE UPPER(feature) = '" & UCASE(p_feature) & "' "

    	set oGetFeatureID = Server.CreateObject("ADODB.Recordset")
   	 oGetFeatureID.Open sSql, Application("DSN"), 3, 1

     if not oGetFeatureID.eof then
        lcl_return = oGetFeatureID("featureid")
     end if

     oGetFeatureID.close
     set oGetFeatureID = nothing
  end if

  getFeatureID = lcl_return
end function

'------------------------------------------------------------------------------
function getCommentsFormID(p_orgid, p_featureid, p_feature)
  lcl_return    = 0
  lcl_featureid = 0

  if p_featureid <> "" then
     lcl_featureid = p_featureid
  else
     if p_feature <> "" then
        lcl_featureid = getFeatureID(p_feature)
     end if
  end if

  sSQL = "SELECT CL_postcomments_formid "
  sSQL = sSQL & " FROM egov_organizations_to_features "
  sSQL = sSQL & " WHERE orgid = "   & p_orgid
  sSQL = sSQL & " AND featureid = " & lcl_featureid

  set oCLFormID = Server.CreateObject("ADODB.Recordset")
  oCLFormID.Open sSQL, Application("DSN"), 3, 1

  if not oCLFormID.eof then
     lcl_return = oCLFormID("CL_postcomments_formid")
  end if

  oCLFormID.close
  set oCLFormID = nothing

  getCommentsFormID = lcl_return

end function

'------------------------------------------------------------------------------
function getCommentsLabel(p_orgid, p_featureid, p_feature)
  lcl_return    = ""
  lcl_featureid = 0

  if p_featureid <> "" then
     lcl_featureid = p_featureid
  else
     if p_feature <> "" then
        lcl_featureid = getFeatureID(p_feature)
     end if
  end if

  sSQL = "SELECT CL_postcomments_label "
  sSQL = sSQL & " FROM egov_organizations_to_features "
  sSQL = sSQL & " WHERE orgid = "   & p_orgid
  sSQL = sSQL & " AND featureid = " & lcl_featureid

  set oCLLabel = Server.CreateObject("ADODB.Recordset")
  oCLLabel.Open sSQL, Application("DSN"), 3, 1

  if not oCLLabel.eof then
     lcl_return = oCLLabel("CL_postcomments_label")
  end if

  oCLLabel.close
  set oCLLabel = nothing

  getCommentsLabel = lcl_return

end function

'------------------------------------------------------------------------------
' getUserEmail( iUserID )
'------------------------------------------------------------------------------
Function getUserEmail( ByVal iUserID )
	Dim sSql, oRs

	getUserEmail = ""

	If iUserID <> "" Then 
		sSql = "SELECT email FROM users WHERE userid = " & iUserID

		Set oRs = Server.CreateObject("ADODB.Recordset")
		oRs.Open sSql, Application("DSN"), 0, 1

		If Not oRs.eof Then 
			getUserEmail = oRs("email")
		End If 

		oRs.Close
		Set oRs = Nothing 
	End If 

End Function 


'------------------------------------------------------------------------------
' RunIdentityInsertStatement( sInsertStatement )
'------------------------------------------------------------------------------
Function RunIdentityInsertStatement( ByVal sInsertStatement )
	Dim sSql, iReturnValue, oInsert

	iReturnValue = 0

'	response.write "<p>" & sInsertStatement & "</p><br /><br />"
'	response.flush

	'INSERT NEW ROW INTO DATABASE AND GET ROWID
	sSql = "SET NOCOUNT ON;" & sInsertStatement & ";SELECT @@IDENTITY AS ROWID;"

	Set oInsert = Server.CreateObject("ADODB.Recordset")
	oInsert.CursorLocation = 3
	oInsert.Open sSql, Application("DSN"), 3, 3
	iReturnValue = oInsert("ROWID")
	oInsert.Close
	Set oInsert = Nothing

	RunIdentityInsertStatement = iReturnValue

End Function 


'------------------------------------------------------------------------------
sub displaySendToOption(iSendToOption, iScreenMode, iIsOnNewLine, iOrgFeature, iUserPermission)

  if iSendToOption <> "" then
    'Check the org feature and user permissions
     if iOrgFeature OR iUserPermission then

       'The screen mode determines if the checkbox(es) are defaulted to "checked" or not.  Screen modes: ADD, EDIT
       'On a "ADD" screen the checkbox(es) are defaulted to "CHECKED"
       'On an "EDIT" screen the checkbox(es) are defaulted to "UNCHECKED"
        lcl_screenmode  = "EDIT"
        lcl_checked     = ""
        lcl_isOnNewLine = "Y"

        if iScreenMode <> "" then
           lcl_screenmode = UCASE(iScreenMode)
        end if

        if lcl_screenmode = "ADD" then
           lcl_checked = " checked=""checked"""
        end if

        if iIsOnNewLine <> "" then
           lcl_isOnNewLine = iIsOnNewLine
        end if

        response.write "<input type=""checkbox"" name=""sendTo_" & iSendToOption & """ id=""sendTo_" & iSendToOption & """ value=""on""" & lcl_checked & " />&nbsp;" & iSendToOption & vbcrlf

       'Check to see if this option will be on it's own "line".
        if iIsOnNewLine = "Y" then
           response.write "<br />" & vbcrlf
        else
           response.write "&nbsp;" & vbcrlf
        end if

     end if
  end if

end sub

'------------------------------------------------------------------------------
function checkSelected( ByVal iValue, ByVal iCompareValue )
  lcl_return = ""

  if iValue <> "" AND iCompareValue <> "" then
     if CStr(iValue) = CStr(iCompareValue) then
        lcl_return = " selected=""selected"""
     end if
  end if

  checkSelected = lcl_return

end function

'------------------------------------------------------------------------------
sub getDatesFromDateRangeChoices( ByVal iDateRange, ByRef lcl_fromDate, ByRef lcl_toDate )

 lcl_today = Date()

 'Today
  if iDateRange = "16" then
     lcl_fromDate = lcl_today
     lcl_toDate   = lcl_today

 'Yesterday
  elseif iDateRange = "17" then
     lcl_fromDate = dateAdd("d",-1,lcl_today)
     lcl_toDate   = dateAdd("d",-1,lcl_today)

 'Tomorrow
  elseif iDateRange = "18" then
     lcl_fromDate = dateAdd("d",+1,lcl_today)
     lcl_toDate   = dateAdd("d",+1,lcl_today)

 'This Month
  elseif iDateRange = "1" then
     lcl_fromDate = month(lcl_today) & "/1/" & year(lcl_today)
     lcl_toDate   = dateAdd("d",-1,dateAdd("m",+1,month(lcl_today) & "/1/" & year(lcl_today)))

 'Next Month
  elseif iDateRange = "13" then
     lcl_nextmonth         = dateadd("m",+2,month(lcl_today) & "/1/" & year(lcl_today))
     lcl_nextmonth_lastday = dateadd("d",-1,lcl_nextmonth)

     lcl_fromDate = month(dateAdd("m",+1,lcl_today)) & "/1/" & year(dateAdd("m",+1,lcl_today))
     lcl_toDate   = month(dateAdd("m",+1,lcl_today)) & "/" & day(lcl_nextmonth_lastday) & "/" & year(dateAdd("m",+1,lcl_today))

 'Last Month
  elseif iDateRange = "2" then
     lcl_fromDate = dateAdd("m",-1,month(lcl_today) & "/1/" & year(lcl_today))
     lcl_toDate   = dateAdd("d",-1,month(lcl_today) & "/1/" & year(lcl_today))

 'This Quarter
  elseif iDateRange = "3" then
     lcl_currentQuarter = getQuarter(month(lcl_today))

     if lcl_currentQuarter = 1 then
        lcl_fromDate = "1/1/"  & year(lcl_today)
        lcl_toDate   = "3/31/" & year(lcl_today)
     elseif lcl_currentQuarter = 2 then
        lcl_fromDate = "4/1/"  & year(lcl_today)
        lcl_toDate   = "6/30/" & year(lcl_today)
     elseif lcl_currentQuarter = 3 then
        lcl_fromDate = "7/1/"  & year(lcl_today)
        lcl_toDate   = "9/30/" & year(lcl_today)
     elseif lcl_currentQuarter = 4 then
        lcl_fromDate = "10/1/"  & year(lcl_today)
        lcl_toDate   = "12/31/" & year(lcl_today)
     end if

 'Next Quarter
  elseif iDateRange = "15" then
     lcl_currentQuarter = getQuarter(month(lcl_today))

     if lcl_currentQuarter = 4 then
        lcl_fromDate = "10/1/"  & year(dateadd("yyyy",+1,lcl_today))
        lcl_toDate   = "12/31/" & year(dateadd("yyyy",+1,lcl_today))
     else
        lcl_pastQuarter = lcl_currentQuarter + 1

        if lcl_pastQuarter = 1 then
           lcl_fromDate = "1/1/"  & year(lcl_today)
           lcl_toDate   = "3/31/" & year(lcl_today)
        elseif lcl_pastQuarter = 2 then
           lcl_fromDate = "4/1/"  & year(lcl_today)
           lcl_toDate   = "6/30/" & year(lcl_today)
        elseif lcl_pastQuarter = 3 then
           lcl_fromDate = "7/1/"  & year(lcl_today)
           lcl_toDate   = "9/30/" & year(lcl_today)
        elseif lcl_pastQuarter = 4 then
           lcl_fromDate = "10/1/"  & year(lcl_today)
           lcl_toDate   = "12/31/" & year(lcl_today)
        end if
     end if

 'Last Quarter
  elseif iDateRange = "4" then
     lcl_currentQuarter = getQuarter(month(lcl_today))

     if lcl_currentQuarter = 1 then
        lcl_fromDate = "10/1/"  & year(dateadd("yyyy",-1,lcl_today))
        lcl_toDate   = "12/31/" & year(dateadd("yyyy",-1,lcl_today))
     else
        lcl_pastQuarter = lcl_currentQuarter - 1

        if lcl_pastQuarter = 1 then
           lcl_fromDate = "1/1/"  & year(lcl_today)
           lcl_toDate   = "3/31/" & year(lcl_today)
        elseif lcl_pastQuarter = 2 then
           lcl_fromDate = "4/1/"  & year(lcl_today)
           lcl_toDate   = "6/30/" & year(lcl_today)
        elseif lcl_pastQuarter = 3 then
           lcl_fromDate = "7/1/"  & year(lcl_today)
           lcl_toDate   = "9/30/" & year(lcl_today)
        elseif lcl_pastQuarter = 4 then
           lcl_fromDate = "10/1/"  & year(lcl_today)
           lcl_toDate   = "12/31/" & year(lcl_today)
        end if
     end if

 'Last Year
  elseif iDateRange = "5" then
     lcl_fromDate = "1/1/"   & year(dateadd("yyyy",-1,lcl_today))
     lcl_toDate   = "12/31/" & year(dateadd("yyyy",-1,lcl_today))

 'Year to Date
  elseif iDateRange = "6" then
     lcl_fromDate = "1/1/" & year(lcl_today)
     lcl_toDate   = lcl_today

 'All Dates to Date
  elseif iDateRange = "7" then
     lcl_fromDate = "1/1/1900"
     lcl_toDate   = lcl_today

 'This Week
  elseif iDateRange = "11" then
     lcl_current_weekday = weekday(lcl_today)
     lcl_weekday_start   = lcl_current_weekday
     lcl_weekday_end     = 7 - lcl_current_weekday

     lcl_fromDate = dateadd("d",-lcl_weekday_start,dateadd("d",+1,lcl_today))
     lcl_toDate   = dateadd("d",+lcl_weekday_end,lcl_today)

 'Next Week
  elseif iDateRange = "14" then
     lcl_current_weekday = weekday(lcl_today)

    'Get the first day of the next week.
     lcl_weekday_start = (7 - lcl_current_weekday) + 1

    'Using the start date simply add 6 days to it and get the end date.
     lcl_weekday_end = lcl_weekday_start + 6

     lcl_fromDate = dateadd("d",+lcl_weekday_start,lcl_today)
     lcl_toDate   = dateadd("d",+lcl_weekday_end,lcl_today)

 'Last Week
  elseif iDateRange = "12" then
     lcl_current_weekday = weekday(lcl_today)

    'First find the end date of the previous week.
     lcl_weekday_end = lcl_current_weekday

    'Using the end date simply take away 6 days to get the start date.
     lcl_weekday_start = lcl_weekday_end+6

     lcl_fromDate = dateadd("d",-lcl_weekday_start,lcl_today)
     lcl_toDate   = dateadd("d",-lcl_weekday_end,lcl_today)

  end if

end sub

'------------------------------------------------------------------------------
function getQuarter( ByVal iMonth )
  lcl_return = 0

  if iMonth <> "" then
     lcl_month = iMonth

     select case lcl_month
        case 1
           lcl_return = 1
        case 2
           lcl_return = 1
        case 3
           lcl_return = 1

        case 4
           lcl_return = 2
        case 5
           lcl_return = 2
        case 6
           lcl_return = 2

        case 7
           lcl_return = 3
        case 8
           lcl_return = 3
        case 9
           lcl_return = 3

        case 10
           lcl_return = 4
        case 11
           lcl_return = 4
        case 12
           lcl_return = 4
     end select
  end if

  getQuarter = lcl_return

end function



'------------------------------------------------------------------------------
sub DrawDateChoices( ByVal sName, ByVal iValue )
	Dim lcl_selected_option0, lcl_selected_option1, lcl_selected_option2, lcl_selected_option4
	Dim lcl_selected_option5, lcl_selected_option6, lcl_selected_option7, lcl_selected_option11
	Dim lcl_selected_option13, lcl_selected_option14, lcl_selected_option15, lcl_selected_option16
	Dim lcl_selected_option17, lcl_selected_option18

	'The following javascript file must be attached.  It contains the "getdates" function used in the OnChange
	'<script language="javascript" src="../scripts/getdates.js"></script>

	lcl_selected_option0  = checkSelected("0",iValue)
	lcl_selected_option1  = checkSelected("1",iValue)
	lcl_selected_option2  = checkSelected("2",iValue)
	lcl_selected_option3  = checkSelected("3",iValue)
	lcl_selected_option4  = checkSelected("4",iValue)
	lcl_selected_option5  = checkSelected("5",iValue)
	lcl_selected_option6  = checkSelected("6",iValue)
	lcl_selected_option7  = checkSelected("7",iValue)
	lcl_selected_option11 = checkSelected("11",iValue)
	lcl_selected_option12 = checkSelected("12",iValue)
	lcl_selected_option13 = checkSelected("13",iValue)
	lcl_selected_option14 = checkSelected("14",iValue)
	lcl_selected_option15 = checkSelected("15",iValue)
	lcl_selected_option16 = checkSelected("16",iValue)
	lcl_selected_option17 = checkSelected("17",iValue)
	lcl_selected_option18 = checkSelected("18",iValue)

	response.write "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """ id=""fromToDateSelection"">" & vbcrlf
	response.write "  <option value=""0"""  & lcl_selected_option0  & ">Or Select Date Range from Dropdown...</option>" & vbcrlf
	response.write "  <option value=""16""" & lcl_selected_option16 & ">Today</option>"             & vbcrlf
	response.write "  <option value=""17""" & lcl_selected_option17 & ">Yesterday</option>"         & vbcrlf
	response.write "  <option value=""18""" & lcl_selected_option18 & ">Tomorrow</option>"          & vbcrlf
	response.write "  <option value=""11""" & lcl_selected_option11 & ">This Week</option>"         & vbcrlf
	response.write "  <option value=""12""" & lcl_selected_option12 & ">Last Week</option>"         & vbcrlf
	response.write "  <option value=""14""" & lcl_selected_option14 & ">Next Week</option>"         & vbcrlf
	response.write "  <option value=""1"""  & lcl_selected_option1  & ">This Month</option>"        & vbcrlf
	response.write "  <option value=""2"""  & lcl_selected_option2  & ">Last Month</option>"        & vbcrlf
	response.write "  <option value=""13""" & lcl_selected_option13 & ">Next Month</option>"        & vbcrlf
	response.write "  <option value=""3"""  & lcl_selected_option3  & ">This Quarter</option>"      & vbcrlf
	response.write "  <option value=""4"""  & lcl_selected_option4  & ">Last Quarter</option>"      & vbcrlf
	response.write "  <option value=""15""" & lcl_selected_option15 & ">Next Quarter</option>"      & vbcrlf
	response.write "  <option value=""6"""  & lcl_selected_option6  & ">Year to Date</option>"      & vbcrlf
	response.write "  <option value=""5"""  & lcl_selected_option5  & ">Last Year</option>"         & vbcrlf
	response.write "  <option value=""7"""  & lcl_selected_option7  & ">All Dates to Date</option>" & vbcrlf
	response.write "</select>" & vbcrlf

End Sub



'------------------------------------------------------------------------------
sub getDelegateInfo( ByVal iUserID, ByRef lcl_delegateid, ByRef lcl_delegate_username, ByRef lcl_delegate_useremail )
  lcl_delegateid         = 0
  lcl_delegate_username  = ""
  lcl_delegate_useremail = ""

  if iUserID <> "" then
     sSql = "SELECT userid, firstname, lastname, email "
     sSql = sSql & " FROM users "
     sSql = sSql & " WHERE userid = (select u2.delegateid "
     sSql = sSql &                 " from users u2 "
     sSql = sSql &                 " where userid = " & iUserID & ") "

     set oDelegateInfo = Server.CreateObject("ADODB.Recordset")
     oDelegateInfo.Open sSql, Application("DSN"), 1, 3

     if not oDelegateInfo.eof then
        lcl_delegateid         = oDelegateInfo("userid")
        lcl_delegate_username  = oDelegateInfo("firstname") & " " & oDelegateInfo("lastname")
        lcl_delegate_useremail = oDelegateInfo("email")
     end if

     oDelegateInfo.close
     set oDelegateInfo = nothing
  end if

end sub

'------------------------------------------------------------------------------
sub setupSendToAndDelegateEmails( ByVal p_sendTo_email, ByVal p_delegate_email, ByRef lcl_email_sendto, ByRef lcl_email_cc )

  lcl_emailaddress_sendto   = ""
  lcl_emailaddress_delegate = ""
  lcl_sendEmailToUser       = "N"
  lcl_sendEmailToDelegate   = "N"
  lcl_email_sendto          = ""
  lcl_email_cc              = ""
  sEmailTo                  = ""
  sEmailDelegate            = ""

 'Setup the SENDTO, CC, and check for a DELEGATE
  if p_sendTo_email <> "" then
     lcl_emailaddress_sendto = formatSendToEmail(p_sendTo_email)

     if isValidEmail(lcl_emailaddress_sendto) then
        sEmailTo            = lcl_emailaddress_sendto
        lcl_sendEmailToUser = "Y"
     end if
  end if

  if p_delegate_email <> "" then
     lcl_emailaddress_delegate = formatSendToEmail(p_delegate_email)

     if isValidEmail(lcl_emailaddress_delegate) then
        sEmailDelegate          = lcl_emailaddress_delegate
        lcl_sendEmailToDelegate = "Y"
     end if
  end if

 'Determine who we are sending the email to.
  if lcl_sendEmailToUser = "Y" AND lcl_sendEmailToDelegate = "Y" then
     lcl_email_sendto = sEmailDelegate
     lcl_email_cc     = sEmailTo
  else
    'If there is no delegate or the delegate's email is invalid
     if lcl_sendEmailToDelegate <> "Y" then

       'The user and cc has a valid email
        if lcl_sendEmailToUser = "Y" then
           lcl_email_sendto = sEmailTo
        end if
     else
        lcl_email_sendto = sEmailDelegate

        if lcl_sendEmailToUser = "Y" then
           lcl_email_cc = sEmailTo
        end if
     end if
  end if
end sub

'------------------------------------------------------------------------------
function getStartingFolder( ByVal iStartFolder )
 lcl_return = "/published_documents"

'Determine which folder to start in
 if trim(iStartFolder) <> "" then
    if ucase(trim(iStartFolder)) = "CITY_ROOT" then
       lcl_return = ""
    else
       lcl_return = "/" & iStartFolder
    end if
 end if

 getStartingFolder = lcl_return

end function




'--------------------------------------------------------------------------------------------------
' GetTimeFormat( dDateTime )
'--------------------------------------------------------------------------------------------------
Function GetTimeFormat( ByVal dDateTime )
	' This takes a datetime and returns the time portion in hh:mm AM/PM
	Dim sHour, sAmPm, sMinute

	sAmPm = "AM"

	sHour = Hour(dDateTime)
	If sHour = 0 Then
		sHour = 12
	ElseIf sHour > 12 Then
		sHour = sHour - 12
		sAmPm = "PM"
	ElseIf sHour = 12 Then
		sAmPm = "PM"
	End If
	
	sMinute = Minute(dDateTime)
	If sMinute < 10 Then
		sMinute = "0" & sMinute
	End If 

	GetTimeFormat = sHour & ":" & sMinute & " " & sAmPm
End Function


'--------------------------------------------------------------------------------------------------
' boolean PaymentGatewayRequiresFeeCheck( iOrgId )
'--------------------------------------------------------------------------------------------------
Function PaymentGatewayRequiresFeeCheck( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT requiresfeecheck FROM egov_payment_gateways G, Organizations O "
	sSql = sSql & "WHERE G.paymentgatewayid = O.OrgPaymentGateway AND O.orgid = " & iOrgId
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("requiresfeecheck") Then
			PaymentGatewayRequiresFeeCheck = True 
		Else
			PaymentGatewayRequiresFeeCheck = False 
		End If 
	Else
		PaymentGatewayRequiresFeeCheck = False 
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function


'------------------------------------------------------------------------------
' double GetProcessingFee( iPaymentId )
'------------------------------------------------------------------------------
Function GetProcessingFee( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(processingfee,0.00) AS processingfee FROM egov_verisign_payment_information "
	sSql = sSql & "WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetProcessingFee = CDbl(oRs("processingfee"))
	Else
		GetProcessingFee = CDbl(0.00)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'------------------------------------------------------------------------------
' string GetProcessingRoute()
'------------------------------------------------------------------------------
Function GetProcessingRoute()
	Dim sSql, oRs, sPaymentGateway

	If session("payment_gateway") <> "" Then
		sPaymentGateway = session("payment_gateway")
	Else
		sPaymentGateway = "0"
	End If 

	sSql = "SELECT ISNULL(processingroute,'') AS processingroute FROM egov_payment_gateways "
	sSQl = sSql & "WHERE paymentgatewayid = " & sPaymentGateway

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetProcessingRoute = oRs("processingroute")
	Else
		GetProcessingRoute = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
' string RemoveAnchorTags( sString )
'------------------------------------------------------------------------------
'Function RemoveAnchorTags( ByVal sString )
'	Dim sNewString, iStart, iEnd

'	sNewString = sString
'	session("Event") = sString

'	If clng(Len(sNewString)) > clng(2) Then 

'		If InStr(sNewString, "<a") > 0 Then 
			' remove all ending anchor tags
'			sNewString = Replace( sNewString, "</a>", "" )

	'		' Try to remove any starting anchor tags
'			Do While InStr(sNewString, "<a") > 0
'				iStart = clng(InStr(sNewString, "<a"))
'				If iStart > clng(1) Then
'					iStart = iStart - clng(1)
'					If clng(InStr(iStart, sNewString,">")) > clng(0) Then 
						' The anchor tag has a complete opening tag
'						iEnd = clng(InStr(iStart, sNewString,">") + 1)
'						sNewString = Mid(sNewString,1,iStart) & Mid(sNewString,iEnd)
'					Else
						' The anchor tag is cut off in the middle of the tag.
'						sNewString = Mid(sNewString,1,iStart)
'					End If 
'				Else
					' THis is for when the string starts with an anchor tag
'					iEnd = clng(InStr(iStart, sNewString,">") + 1)
'					sNewString = Mid(sNewString, iEnd)
'				End If 
'			Loop
			'sNewString = sNewString & " [" & iStart & "] [" & iEnd & "] [" & sNextString & "]"
'		End If 

'	End If
	
'	RemoveAnchorTags = sNewString 
'	session("Event") = ""

'End Function

'------------------------------------------------------------------------------
function RemoveAnchorTags( ByVal sString )
	Dim sNewString, iStart, iEnd

	sNewString       = sString
	session("Event") = sString

	if clng(len(sNewString)) > clng(2) then
    if instr(sNewString,"<a") > 0 then
       do while instr(sNewString,"<a") > 0
          lcl_string_length = len(sNewString)

         'Remove all anchor tags
          lcl_string_length   = len(sNewString)
          lcl_anchortag_start = clng(instr(sNewString,"<a") - 1)
          lcl_anchortag_close = clng(instr(sNewString,"</a>"))
          lcl_left_string     = ""
          lcl_right_string    = ""

         'Break the string up to the "left" side, part BEFORE the start of the opening anchor tag
         'and the "right" side, rest of string STARTING at the opening anchor tag
          lcl_left_string  = left(sNewString,lcl_anchortag_start)

          if  clng(instr(lcl_anchortag_start+1,sNewString,">")) > lcl_anchortag_start _
          AND clng(instr(lcl_anchortag_start+1,sNewString,">")) < lcl_anchortag_close then
              lcl_anchortag_end = clng(instr(lcl_anchortag_start+1,sNewString,">") + 1)
              lcl_right_string  = mid(sNewString,lcl_anchortag_end)
              lcl_right_string  = replace(lcl_right_string,"</a>","",1)
          else
              lcl_left_string = replace(lcl_left_string,"</a>","",1)
          end if

          sNewString = lcl_left_string & lcl_right_string
       loop
    end if
 end if

 RemoveAnchorTags = sNewString
	session("Event") = ""

end function

'------------------------------------------------------------------------------
function checkIsPushForm( ByVal iOrgID, ByVal iFormID )
	Dim sSQL, oGetCalForm

  lcl_return = False

  if iFormID <> "" then

    'Set up the formids that are considered to be "push" forms
     lcl_push_formids = ""

     sSQL = "SELECT OrgRequestCalForm "
     sSQL = sSQL & " FROM organizations "
     sSQL = sSQL & " WHERE orgid = " & iOrgID

    	set oGetCalForm = Server.CreateObject("ADODB.Recordset")
   	 oGetCalForm.Open sSQL, Application("DSN") , 3, 1

     if not oGetCalForm.eof then
        if lcl_push_formids <> "" then
           lcl_push_formids = lcl_push_formids & "," & oGetCalForm("OrgRequestCalForm")
        else
           lcl_push_formids = oGetCalForm("OrgRequestCalForm")
        end if
     end if

     oGetCalForm.close
     set oGetCalForm = nothing

    'Check to see if the formid passed in is considered a "push" form
     if lcl_push_formids <> "" then
        lcl_formid = iFormID

        sSQL = "SELECT 'Y' as lcl_exists "
        sSQL = sSQL & " FROM organizations "
        sSQL = sSQL & " WHERE orgid = " & iOrgID
        sSQL = sSQL & " AND " & lcl_formid & " IN (" & lcl_push_formids & ") "

       	set oIsPushForm = Server.CreateObject("ADODB.Recordset")
      	 oIsPushForm.Open sSQL, Application("DSN") , 3, 1

        if not oIsPushForm.eof then
           lcl_return = True
        end if

        oIsPushForm.close
        set oIsPushForm = nothing
     end if

  end if

  checkIsPushForm = lcl_return

end function

'------------------------------------------------------------------------------
sub formatEventDateTime(ByVal p_date, ByVal p_end_date, ByRef sDate1, ByRef sDate2)

 'Used to trim seconds from dates displayed on calendar pages.
		sDate1 = cStr(p_date)
		sDate2 = cStr(p_end_date)

		iTrimDate1 = clng(InStrRev(sDate1,":"))
		iTrimDate2 = clng(InStrRev(sDate2,":"))

	'Retrieves AM/PM, trims final :00 and builds string
		if iTrimDate1 > 0 then
  			sTemp  = Right(sDate1, 2)
		  	sDate1 = Left(sDate1,iTrimDate1 - 1) & " " & sTemp
				 sTemp  = ""
  end if

		if iTrimDate2 > 0 then
     sTemp  = Right(sDate2, 2)
     sDate2 = Left(sDate2,iTrimDate2 - 1) & " " & sTemp
	    sTemp  = ""
  end if

end sub

'------------------------------------------------------------------------------
function containsApostrophe(p_value)

  lcl_return = False

  if p_value <> "" then
     lcl_value = p_value

     if isnumeric(lcl_value) then
        lcl_value = CStr(lcl_value)
     end if

     if instr(lcl_value,"'") > 0 then
        lcl_return = True
     end if
  end if

  containsApostrophe = lcl_return

end function

'------------------------------------------------------------------------------
function getGoogleSearchID(iOrgID, iDBColumn)
	 dim sSQL, sDBColumn, sOrgID, lcl_return

  sOrgID     = 0
  sDBColumn  = ""
  lcl_return = ""

  if not containsApostrophe(iOrgID) then
     sOrgID = clng(iOrgID)
  end if

  if not containsApostrophe(iDBColumn) then
     sDBColumn = iDBColumn
  end if

  sSQL = "SELECT " & sDBColumn & " as googleSearchID "
  sSQL = sSQL & " FROM organizations "
  sSQL = sSQL & " WHERE orgid = " & sOrgID

 	set oRs = Server.CreateObject("ADODB.Recordset")
 	oRs.Open sSQL, Application("DSN"), 3, 1

 	if not oRs.eof then
	   	lcl_return = oRs("googleSearchID")
  end if

 	oRs.Close
  set oRs = nothing

  getGoogleSearchID = lcl_return

end function

'------------------------------------------------------------------------------
sub displayPrivacyPolicyFields(iScreenType, iFieldType, iValue)

  dim lcl_screentype, lcl_fieldlabel, lcl_fieldid, lcl_fieldtype, lcl_fieldvalue

  lcl_screentype   = "EDIT"
  lcl_fieldid      = "egov"
  lcl_fieldtype    = "EGOV"
  lcl_fieldlabel   = "Footer - E-Gov"
  lcl_fieldvalue   = ""
  lcl_website_url  = ""
  lcl_website_text = ""
  lcl_display_text = ""
  lcl_url_start    = 0
  lcl_url_end      = 0
  lcl_url_length   = 0
  lcl_text_start   = 0
  lcl_text_end     = 0
  lcl_text_length  = 0

  if iScreenType <> "" then
     lcl_screentype = ucase(iScreenType)
  end if

  if iFieldType <> "" then
     lcl_fieldtype = ucase(iFieldType)
  end if

  if lcl_fieldtype = "MOBILE" then
     lcl_fieldid    = "mobile"
     lcl_fieldlabel = "Footer - Mobile"
  end if

  if lcl_screentype = "EDIT" then
     if iValue <> "" then
        lcl_fieldvalue = iValue

        lcl_url_start      = instr(lcl_fieldvalue,"[")
        lcl_url_end        = instr(lcl_fieldvalue,"]")
        lcl_url_length     = lcl_url_end - lcl_url_start

        lcl_text_start     = instr(lcl_fieldvalue,"<")
        lcl_text_end       = instr(lcl_fieldvalue,">")
        lcl_text_length    = lcl_text_end - lcl_text_start

        lcl_website_url  = mid(lcl_fieldvalue,lcl_url_start,lcl_url_length)
        lcl_website_url  = replace(lcl_website_url,"[","")
        lcl_website_url  = replace(lcl_website_url,"]","")

        lcl_website_text = mid(lcl_fieldvalue,lcl_text_start,lcl_text_length)
        lcl_website_text = replace(lcl_website_text,"<","")
        lcl_website_text = replace(lcl_website_text,">","")

        lcl_display_text = lcl_website_url
        lcl_url_value    = replace(lcl_url_value,lcl_fieldvalue,"")

        if lcl_website_text <> "" then
           lcl_display_text = lcl_website_text
        end if
     end if
  else
     if lcl_fieldtype = "EGOV" then
        lcl_website_url  = "privacy_policy.asp"
        lcl_display_text = "Privacy Policy"
     end if
  end if

  response.write "<strong>" & lcl_fieldlabel & ":</strong><br />" & vbcrlf
  response.write "URL: <input type=""text"" name=""privacypolicy_url_" & lcl_fieldid & """ id=""privacypolicy_url_" & lcl_fieldid & """ value=""" & lcl_website_url & """ size=""40"" maxlength=""1000"" />" & vbcrlf
  response.write "Display Value: <input type=""text"" name=""privacypolicy_text_" & lcl_fieldid & """ id=""privacypolicy_text_" & lcl_fieldid & """ value=""" & lcl_display_text & """ size=""30"" maxlength=""1000"" />" & vbcrlf
  response.write "<br />" & vbcrlf
  response.write "<input type=""hidden"" name=""privacypolicy_" & lcl_fieldid & """ id=""privacypolicy_" & lcl_fieldid & """ value=""" & lcl_fieldvalue & """ size=""50"" maxlength=""4000"" />" & vbcrlf

  if lcl_fieldtype = "EGOV" then
     response.write "<br /><br />" & vbcrlf
  end if

end sub



'--------------------------------------------------------------------------------------------------
' void SetUnDoBtnDisplay iPaymentId, bShowUnDoBtn
'--------------------------------------------------------------------------------------------------
Sub SetUnDoBtnDisplay( ByVal iPaymentId, ByVal bShowUnDoBtn )
	Dim sSql, oRs, showUnDoBtnValue
	
	if bShowUnDoBtn then 
		showUnDoBtnValue = "1"
	else
		showUnDoBtnValue = "0"
	end if 

	sSql = "UPDATE egov_class_payment SET showundobtn = " & showUnDoBtnValue & " WHERE paymentid = " & iPaymentId
	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean IsUnDoBtnDisplayed( iPaymentId )
'--------------------------------------------------------------------------------------------------
Function IsUnDoBtnDisplayed( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT showundobtn FROM egov_class_payment WHERE paymentid = " & iPaymentId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		IsUnDoBtnDisplayed = oRs("showundobtn")
	Else
		IsUnDoBtnDisplayed = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

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
	    sURL = oRs("OrgPublicWebsiteURL") & "citizen-action-line/"
    	elseif sColumnName = "wp_subscriptions_url" then
	    sURL = oRs("OrgPublicWebsiteURL") & "subscriptions/"
    	end if
    end if
  End If

  'sURL = replace(sURL,"http://glenford.egovhost2.com","https://www2.egovlink.com")
  
  oRs.Close
  Set oRs = Nothing  
  
  getOrganization_WP_URL = sURL
   
End Function

Function AuditCookie(intTWFUserid)
	AuditCookie = False
	sSQL = "SELECT auditdatetime FROM auditlog WHERE userid = '" & dbsafe(intTWFUserid) & "' AND auditnotes = '" & Request.ServerVariables("REMOTE_ADDR") & "' AND auditdatetime >= '" & DateAdd("d", -60, now()) & "'"
	Set oAudit = Server.CreateObject("ADODB.RecordSet")
	oAudit.Open sSQL, Application("DSN"), 3, 1
	if not oAudit.EOF then AuditCookie = true
	oAudit.Close
	Set oAudit = Nothing
	
End Function

Function hasWordPress()
	hasWordPress = false
	sSQL = "SELECT wpLive FROM organizations WHERE wpLive = 1 AND orgid = " & session("orgid")
	Set oWP = Server.CreateObject("ADODB.RecordSet")
	oWP.Open sSQL, Application("DSN"), 3, 1
	if not oWP.EOF then hasWordPress = true
	oWP.Close
	Set oWP = Nothing


End Function

%>
